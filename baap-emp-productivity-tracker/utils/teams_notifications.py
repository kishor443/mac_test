"""
Microsoft Teams notifications utility
Gets Teams messages/notifications from Windows notification system or Teams API
"""
import os
import json
import sqlite3
import threading
import time
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Optional

# Global cache to track previous window states and detect new messages
_previous_teams_windows = {}  # {hwnd: (sender, message, timestamp)}
_last_message_check_time = None

try:
    import win32gui
    import win32process
    import psutil
except ImportError:
    win32gui = None
    win32process = None
    psutil = None

from utils.logger import logger


def get_teams_notification_db_path() -> Optional[Path]:
    """Get path to Teams notification database"""
    try:
        appdata = os.getenv('APPDATA')
        if not appdata:
            return None
        
        # Teams stores notifications in different locations
        teams_paths = [
            Path(appdata) / 'Microsoft' / 'Teams' / 'IndexedDB' / 'https_teams.microsoft.com_0.indexeddb.leveldb',
            Path(appdata) / 'Microsoft' / 'Teams' / 'storage' / 'default' / 'indexeddb',
            Path(os.getenv('LOCALAPPDATA', '')) / 'Microsoft' / 'Teams' / 'IndexedDB',
        ]
        
        for path in teams_paths:
            if path.exists():
                return path.parent
        return None
    except Exception as e:
        logger.error(f"Error getting Teams DB path: {e}")
        return None


# Global cache to track previous window states and detect new messages
_previous_teams_windows = {}  # {hwnd: (sender, message, timestamp)}
_last_message_check_time = None

def get_teams_messages_from_window() -> List[Dict]:
    """Extract Teams messages from active Teams window titles - only NEW messages"""
    global _previous_teams_windows, _last_message_check_time
    messages = []
    try:
        if not win32gui:
            return messages
        
        current_time = time.time()
        current_windows = {}  # Track current state
        
        def enum_windows_callback(hwnd, results):
            try:
                window_title = win32gui.GetWindowText(hwnd)
                if not window_title:
                    return True
                
                # Teams window titles often have format: "Chat | Channel/Person Name | Message preview"
                # or "Microsoft Teams - Chat | ..."
                title_lower = window_title.lower()
                # Check for Teams-related windows more broadly
                # Teams windows can have various formats:
                # - "Microsoft Teams"
                # - "Chat | Person Name | Message"
                # - "Chat | Channel Name | Message"
                # - "Personal | Person Name | Message"
                is_teams_window = (
                    'teams' in title_lower or 
                    ('chat' in title_lower and ' | ' in window_title) or
                    ('personal' in title_lower and ' | ' in window_title) or
                    ('conversation' in title_lower and ' | ' in window_title)
                )
                
                if is_teams_window:
                    logger.info(f"Found Teams window: {window_title}")
                    # Parse Teams window title to extract message info
                    # Format examples:
                    # "Chat | PRODUCTIVITY-TRACKER_DESIGNING-TEAM | P..."
                    # "Chat | John Doe | Hello, how are you?"
                    # "Microsoft Teams - Chat | Channel Name | Message preview"
                    # "Personal | Person Name | Message"
                    
                    # Remove "Microsoft Teams -" prefix if present
                    clean_title = window_title
                    if ' - ' in clean_title:
                        parts = clean_title.split(' - ', 1)
                        if len(parts) > 1:
                            clean_title = parts[1]
                    
                    # Split by pipe to get components
                    if ' | ' in clean_title:
                        parts = clean_title.split(' | ')
                        logger.info(f"Window title parts: {parts}")
                        
                        if len(parts) >= 2:
                            chat_type = parts[0].strip()  # "Chat", "Personal", or "Conversation"
                            sender_or_channel = parts[1].strip()  # Person name or channel name
                            
                            logger.info(f"Extracted sender: {sender_or_channel}, chat_type: {chat_type}")
                            
                            # Extract message preview if available
                            message_preview = ''
                            if len(parts) >= 3:
                                message_preview = parts[2].strip()
                            
                            # If no message preview, try to extract from title differently
                            if not message_preview and len(parts) == 2:
                                # Sometimes format is: "Chat | Person Name: Message preview"
                                if ':' in sender_or_channel:
                                    sender_parts = sender_or_channel.split(':', 1)
                                    sender_or_channel = sender_parts[0].strip()
                                    message_preview = sender_parts[1].strip() if len(sender_parts) > 1 else ''
                            
                            # Clean up message preview (remove trailing dots if truncated)
                            if message_preview.endswith('...') or message_preview.endswith('P...'):
                                message_preview = message_preview.rstrip('.')
                            
                            # If still no message, check if last part might be message
                            if not message_preview:
                                # Sometimes the format is different - check all parts
                                for part in parts[2:]:
                                    if part.strip() and part.strip() not in ['Personal', 'Chat', 'Conversation']:
                                        message_preview = part.strip()
                                        break
                            
                            # If still no message, use a default based on chat type
                            if not message_preview:
                                if 'personal' in chat_type.lower() or 'personal' in sender_or_channel.lower():
                                    message_preview = 'Personal chat'
                                else:
                                    message_preview = 'New message'
                            
                            # Generate unique timestamp for each detection
                            current_time_str = datetime.now().strftime('%I:%M %p')
                            
                            # Store current window state (always track)
                            current_windows[hwnd] = (sender_or_channel, message_preview, current_time)
                            
                            # Check if this is a NEW message (different from previous state)
                            is_new_message = False
                            if hwnd in _previous_teams_windows:
                                prev_sender, prev_message, prev_time = _previous_teams_windows[hwnd]
                                # It's new if message content changed (actual new message)
                                if message_preview and message_preview != prev_message:
                                    is_new_message = True
                            else:
                                # First time seeing this window - consider it new if it has message content
                                if message_preview:
                                    is_new_message = True
                            
                            # Add ALL messages (not just new ones) - we'll filter in UI
                            # But mark which ones are new for highlighting
                            results.append({
                                'sender': sender_or_channel,
                                'message': message_preview,
                                'time': current_time_str,
                                'source': 'Teams Window',
                                'chat_type': chat_type,
                                'window_title': window_title,  # Keep original for debugging
                                'is_new': is_new_message
                            })
                    else:
                        # If no pipe separator, try to extract from title directly
                        # Look for patterns like "Chat with John" or "Channel Name"
                        if 'chat' in title_lower or 'conversation' in title_lower:
                            # Extract name after "Chat" or "Chat with"
                            if 'chat with' in title_lower:
                                name_start = title_lower.find('chat with') + len('chat with')
                                name = clean_title[name_start:].strip()
                                if name:
                                    results.append({
                                        'sender': name,
                                        'message': 'Active chat',
                                        'time': datetime.now().strftime('%I:%M %p'),
                                        'source': 'Teams Window'
                                    })
            except Exception as e:
                logger.debug(f"Error parsing Teams window: {e}")
            return True
        
        windows = []
        win32gui.EnumWindows(enum_windows_callback, windows)
        
        # Debug: log found windows
        if windows:
            logger.debug(f"Found {len(windows)} Teams windows")
        
        # Update previous windows state for next check
        _previous_teams_windows = current_windows.copy()
        _last_message_check_time = current_time
        
        # Remove duplicates - group by sender, keep most recent message
        sender_messages = {}  # {sender: most_recent_message}
        
        for msg in windows:
            sender = msg.get('sender', '').strip()
            message = msg.get('message', '').strip()
            
            if not sender:  # Skip if no sender
                continue
            
            # Keep the most recent message for each sender (both new and existing)
            if sender not in sender_messages:
                sender_messages[sender] = msg
            else:
                # Compare timestamps - keep the newer one
                existing_time = sender_messages[sender].get('time', '')
                new_time = msg.get('time', '')
                # Also prefer new messages over old ones
                if msg.get('is_new', False) or new_time > existing_time:
                    sender_messages[sender] = msg
            
            # Add chat identifier for opening Teams
            msg['chat_identifier'] = sender
        
        # Convert to list
        unique_messages = list(sender_messages.values())
        
        # Sort by time (newest first) - new messages first
        unique_messages.sort(key=lambda x: (
            not x.get('is_new', False),  # New messages first
            x.get('time', '')
        ), reverse=True)
        
        # Debug: log final messages
        if unique_messages:
            logger.debug(f"Returning {len(unique_messages)} unique Teams messages")
        
        # Return all unique messages (no limit)
        return unique_messages
    except Exception as e:
        logger.error(f"Error getting Teams messages from window: {e}")
        return messages


def get_teams_notifications_from_system() -> List[Dict]:
    """Get Teams notifications from Windows notification system"""
    messages = []
    try:
        # Try to read from Windows notification database
        # This is a simplified approach - actual implementation would need
        # to parse Windows notification database
        appdata = os.getenv('LOCALAPPDATA')
        if appdata:
            notification_path = Path(appdata) / 'Packages' / 'Microsoft.Windows.ShellExperienceHost_cw5n1h2txyewy' / 'AC' / 'INetCache'
            # This is a placeholder - actual notification DB location varies
            pass
        
        # For now, return empty list
        # In production, you'd parse the actual notification database
        return messages
    except Exception as e:
        logger.error(f"Error getting Teams notifications from system: {e}")
        return messages


def get_teams_messages_simple() -> List[Dict]:
    """Simple method to get Teams messages - checks if Teams is running and returns sample data"""
    messages = []
    try:
        # Check if Teams is running
        teams_running = False
        if psutil:
            for proc in psutil.process_iter(['name']):
                try:
                    proc_name = proc.info['name'].lower()
                    if 'teams' in proc_name:
                        teams_running = True
                        break
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue
        
        if teams_running:
            # Try to get messages from window
            window_messages = get_teams_messages_from_window()
            if window_messages:
                return window_messages
            
            # If no window messages, return sample/placeholder
            # In production, you'd integrate with Teams Graph API
            return [
                {
                    'sender': 'Teams Notification',
                    'message': 'New message in Teams',
                    'time': datetime.now().strftime('%I:%M %p'),
                    'source': 'Teams'
                }
            ]
        
        return messages
    except Exception as e:
        logger.error(f"Error in get_teams_messages_simple: {e}")
        return messages


def get_teams_messages_via_graph_api(access_token: Optional[str] = None) -> List[Dict]:
    """Get Teams messages via Microsoft Graph API (works without opening Teams)"""
    messages = []
    try:
        if not access_token:
            logger.debug("No access token available for Graph API")
            return messages
        
        import requests
        
        # Get recent chat messages
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }
        
        logger.debug("Fetching Teams messages via Graph API...")
        
        # Get chats
        chats_url = 'https://graph.microsoft.com/v1.0/me/chats'
        response = requests.get(chats_url, headers=headers, timeout=10)
        
        # Suppress 401 errors to avoid log spam (invalid token)
        if response.status_code == 401:
            logger.debug("Graph API authentication failed - token invalid or expired. Falling back to window detection.")
            return messages
        
        if response.status_code == 200:
            chats_data = response.json()
            chats = chats_data.get('value', [])
            
            logger.info(f"Found {len(chats)} chats via Graph API")
            
            # Get messages from recent chats
            for chat in chats[:10]:  # Increased to 10 recent chats
                chat_id = chat.get('id')
                chat_type = chat.get('chatType', '')
                
                if not chat_id:
                    continue
                
                # Get messages for this chat
                messages_url = f'https://graph.microsoft.com/v1.0/me/chats/{chat_id}/messages?$top=1'
                msg_response = requests.get(messages_url, headers=headers, timeout=5)
                
                if msg_response.status_code == 200:
                    msg_data = msg_response.json()
                    recent_messages = msg_data.get('value', [])
                    
                    for msg in recent_messages:
                        from_data = msg.get('from', {})
                        user_data = from_data.get('user', {}) if from_data else {}
                        sender = user_data.get('displayName', 'Unknown') if user_data else 'Unknown'
                        
                        body_data = msg.get('body', {})
                        body = body_data.get('content', '') if body_data else ''
                        
                        # Strip HTML tags from message
                        import re
                        body = re.sub('<[^<]+?>', '', body)
                        
                        created = msg.get('createdDateTime', '')
                        
                        # Parse time
                        try:
                            from dateutil import parser
                            dt = parser.isoparse(created)
                            time_str = dt.strftime('%I:%M %p')
                        except Exception:
                            try:
                                dt = datetime.fromisoformat(created.replace('Z', '+00:00'))
                                time_str = dt.strftime('%I:%M %p')
                            except Exception:
                                time_str = datetime.now().strftime('%I:%M %p')
                        
                        # Full message for notifications
                        full_message = body
                        
                        # Truncate for display
                        if len(body) > 100:
                            body = body[:97] + '...'
                        
                        if sender and sender != 'Unknown':
                            messages.append({
                                'sender': sender,
                                'message': body,
                                'full_message': full_message,
                                'time': time_str,
                                'source': 'Teams Graph API',
                                'chat_id': chat_id,
                                'chat_identifier': sender,
                                'chat_type': chat_type,
                                'is_new': True  # Mark as new for initial fetch
                            })
            
            logger.debug(f"Extracted {len(messages)} messages from Graph API")
        elif response.status_code != 401:  # Don't log 401 errors (already handled above)
            logger.debug(f"Graph API returned status {response.status_code}")
        
        # Return all messages
        return messages
    except Exception as e:
        logger.error(f"Error getting Teams messages via Graph API: {e}")
        return messages


def get_teams_messages(access_token: Optional[str] = None) -> List[Dict]:
    """
    Main function to get Teams messages (works even without opening Teams if API access available)
    Tries multiple methods: Graph API > Window detection > Simple check
    """
    messages = []
    all_messages = []
    
    # Try Graph API first (if token available) - works WITHOUT opening Teams
    if access_token:
        try:
            logger.debug("Attempting to fetch Teams messages via Graph API...")
            api_messages = get_teams_messages_via_graph_api(access_token)
            if api_messages:
                logger.debug(f"Graph API returned {len(api_messages)} messages")
                all_messages.extend(api_messages)
        except Exception as e:
            logger.debug(f"Graph API method failed: {e}")
    else:
        logger.debug("No access token available, skipping Graph API")
    
    # Try window detection - gets real-time window titles (requires Teams open)
    try:
        window_messages = get_teams_messages_from_window()
        if window_messages:
            logger.debug(f"Window detection returned {len(window_messages)} messages")
            # Add window messages, avoiding duplicates
            for wmsg in window_messages:
                # Check if not already in all_messages
                is_duplicate = any(
                    m.get('sender') == wmsg.get('sender') and 
                    m.get('message', '')[:50] == wmsg.get('message', '')[:50]
                    for m in all_messages
                )
                if not is_duplicate:
                    all_messages.append(wmsg)
    except Exception as e:
        logger.debug(f"Window detection method failed: {e}")
    
    # If we have messages from either source, return them
    if all_messages:
        # Sort by time (newest first)
        all_messages.sort(key=lambda x: x.get('time', ''), reverse=True)
        # Remove duplicates by sender
        seen_senders = set()
        unique_messages = []
        for msg in all_messages:
            sender = msg.get('sender')
            if sender and sender not in seen_senders:
                seen_senders.add(sender)
                unique_messages.append(msg)
        logger.debug(f"Returning {len(unique_messages)} unique messages")
        return unique_messages
    
    # Fallback to simple method
    try:
        simple_messages = get_teams_messages_simple()
        if simple_messages:
            logger.debug("Using simple fallback method")
            return simple_messages
    except Exception as e:
        logger.debug(f"Simple method failed: {e}")
    
    logger.debug("No Teams messages found from any source")
    return messages

