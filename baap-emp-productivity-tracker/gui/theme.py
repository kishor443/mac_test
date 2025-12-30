import tkinter as tk
from tkinter import ttk

# Centralized spacing and font tokens for consistent UI
spacing = {
	"xs": 2,
	"sm": 4,
	"md": 8,
	"lg": 12,
	"xl": 16,
}

fonts = {
	"title": ("Figtree", 12, "bold"),
	"subtitle": ("Figtree", 10, "bold"),
	"body": ("Figtree", 10),
}


def apply_theme(root: tk.Misc) -> None:
	"""Apply a light professional ttk theme with consistent styles.

	This sets a base theme (clam) and configures common widget styles.
	"""
	style = ttk.Style(master=root)
	# Use a solid, widely available theme as base
	try:
		style.theme_use("clam")
	except tk.TclError:
		# Fall back gracefully
		available = style.theme_names()
		if available:
			style.theme_use(available[0])

	# Palette
	# Palette - Modern Pastel / Glassmorphism inspired
	bg = "#F3F4F6"       # Light gray/white background
	fg = "#1F2937"       # Dark gray text
	muted = "#6B7280"    # Muted gray
	primary = "#8B5CF6"  # Soft Violet
	primary_hover = "#7C3AED"
	secondary = "#F472B6" # Soft Pink
	secondary_hover = "#EC4899"
	success = "#34D399"  # Soft Green
	success_hover = "#10B981"
	warning = "#FBBF24"  # Soft Amber
	warning_hover = "#F59E0B"
	danger = "#F87171"   # Soft Red
	danger_hover = "#EF4444"
	info = "#60A5FA"     # Soft Blue
	info_hover = "#3B82F6"
	card_bg = "#FFFFFF"
	alt_bg = "#E5E7EB"
	border = "#E5E7EB"

	# Window background
	if isinstance(root, (tk.Tk, tk.Toplevel)):
		root.configure(background=bg)

	# Base styles
	style.configure("TFrame", background=bg)
	style.configure("TLabel", background=bg, foreground=fg, font=fonts["body"])
	# Glass-like Button Styles
	# Base TButton: White with light border, hover effect
	style.configure("TButton", font=fonts["body"], padding=(10, 8), borderwidth=1, relief="raised", background="#FFFFFF", bordercolor="#E5E7EB", focuscolor=primary)

	style.map(
		"TButton",
		foreground=[("disabled", muted), ("active", fg)],
		background=[("active", "#F9FAFB"), ("!active", "#FFFFFF")],
		bordercolor=[("active", primary)],
		relief=[("pressed", "sunken"), ("!pressed", "raised")],
	)
	
	# Primary: Solid color but with a "glassy" border highlight
	style.configure(
		"Primary.TButton",
		background=primary,
		foreground="#FFFFFF",
		bordercolor="#A78BFA", # Lighter violet border
		focuscolor=primary,
		borderwidth=1,
		relief="solid"
	)
	style.map(
		"Primary.TButton",
		background=[("active", primary_hover), ("!active", primary)],
		foreground=[("!disabled", "#FFFFFF")],
		bordercolor=[("active", "#C4B5FD")], # Even lighter on hover
	)

	# Secondary/Accent buttons
	style.configure("Secondary.TButton", background=secondary, foreground=fg, bordercolor=secondary, focuscolor=secondary)
	style.map("Secondary.TButton", background=[("active", secondary_hover), ("!active", secondary)], foreground=[("!disabled", fg)])
	style.configure("Success.TButton", background=success, foreground=fg, bordercolor=success, focuscolor=success)
	style.map("Success.TButton", background=[("active", success_hover), ("!active", success)], foreground=[("!disabled", fg)])
	style.configure("Warning.TButton", background=warning, foreground=fg, bordercolor=warning, focuscolor=warning)
	style.map("Warning.TButton", background=[("active", warning_hover), ("!active", warning)], foreground=[("!disabled", fg)])
	style.configure("Danger.TButton", background=danger, foreground=fg, bordercolor=danger, focuscolor=danger)
	style.map("Danger.TButton", background=[("active", danger_hover), ("!active", danger)], foreground=[("!disabled", fg)])
	style.configure("Info.TButton", background=info, foreground=fg, bordercolor=info, focuscolor=info)
	style.map("Info.TButton", background=[("active", info_hover), ("!active", info)], foreground=[("!disabled", fg)])

	# Custom Colors for UI Match
	# Red Button (Clock In)
	red_btn = "#F87171"
	red_btn_hover = "#EF4444"
	style.configure("Red.TButton", background=red_btn, foreground="#000000", bordercolor=red_btn, focuscolor=red_btn, font=("Figtree", 9, "bold"))
	style.map("Red.TButton", background=[("active", red_btn_hover), ("!active", red_btn)], foreground=[("!disabled", "#000000")])

	# Blue Button (End Break)
	blue_btn = "#60A5FA"
	blue_btn_hover = "#3B82F6"
	style.configure("Blue.TButton", background=blue_btn, foreground="#000000", bordercolor=blue_btn, focuscolor=blue_btn, font=("Figtree", 9, "bold"), relief="raised", borderwidth=2)
	style.map("Blue.TButton", background=[("active", blue_btn_hover), ("!active", blue_btn)], foreground=[("!disabled", "#000000")], relief=[("pressed", "sunken"), ("!pressed", "raised")])

	# Purple Button (Refresh Attendance)
	purple_btn = "#8B5CF6"
	purple_btn_hover = "#7C3AED"
	style.configure("Purple.TButton", background=purple_btn, foreground="#FFFFFF", bordercolor=purple_btn, focuscolor=purple_btn)
	style.map("Purple.TButton", background=[("active", purple_btn_hover), ("!active", purple_btn)], foreground=[("!disabled", "#FFFFFF")])


	# Header label
	style.configure("Header.TLabel", font=fonts["title"], foreground=fg)
	style.configure("Subheader.TLabel", font=fonts["subtitle"], foreground=muted)
	style.configure("Accent.TLabel", font=fonts["subtitle"], foreground=primary)

	# Status bar
	style.configure(
		"Status.TLabel",
		background=alt_bg,
		foreground=muted,
		font=fonts["body"],
	)

	# Notebook (tabs)
	style.configure("TNotebook", background=bg, borderwidth=0)
	style.configure("TNotebook.Tab", padding=(spacing["md"], spacing["sm"]), font=fonts["body"])
	style.map("TNotebook.Tab",
		background=[("selected", card_bg), ("!selected", bg)],
		foreground=[("selected", fg), ("!selected", muted)],
	)

	# Scrollbar
	style.configure("Vertical.TScrollbar", gripcount=0,
		background=alt_bg, darkcolor=bg, lightcolor=bg,
		troughcolor=bg, bordercolor=bg, arrowcolor=muted,
		arrowsize=12, width=10) # Reduced width


	# Entry / Combobox
	style.configure("TEntry", padding=(8, 6))
	try:
		style.configure("TEntry", fieldbackground=card_bg)
	except tk.TclError:
		pass
	style.configure("TCombobox", padding=(8, 6))

	# Labelframe
	style.configure("TLabelframe", background=bg, bordercolor=border)
	style.configure("TLabelframe.Label", background=bg, foreground=muted, font=fonts["subtitle"])

	# Progressbar
	style.configure("Primary.Horizontal.TProgressbar", troughcolor=alt_bg, background=primary)
	style.configure("Success.Horizontal.TProgressbar", troughcolor=alt_bg, background=success)


