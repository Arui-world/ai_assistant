import frappe


def ensure_ai_assistant_sidebar_visible():
	"""Keep AI Assistant visible for non-admin Desk users.

	Frappe filters Workspace Sidebar records by the user's allowed modules. The
	ai_assistant app has a Page but no readable DocType, so normal users do not
	get the AI Assistant module in allow_modules. Leaving the sidebar module blank
	makes the public Page link eligible for all users who can access ai-chat.
	"""
	if frappe.db.exists("Workspace Sidebar", "AI Assistant"):
		frappe.db.set_value(
			"Workspace Sidebar",
			"AI Assistant",
			{
				"app": "ai_assistant",
				"module": "",
				"standard": 1,
			},
			update_modified=False,
		)

	if frappe.db.exists("Desktop Icon", "AI Assistant"):
		frappe.db.set_value(
			"Desktop Icon",
			"AI Assistant",
			{
				"app": "ai_assistant",
				"hidden": 0,
				"link_to": "AI Assistant",
				"link_type": "Workspace Sidebar",
				"standard": 1,
				"idx": 0,
			},
			update_modified=False,
		)

	if frappe.db.exists("Page", "ai-chat"):
		page = frappe.get_doc("Page", "ai-chat")
		if "All" not in {row.role for row in page.roles}:
			page.append("roles", {"role": "All"})
			page.flags.ignore_permissions = True
			page.save(ignore_permissions=True)

	frappe.cache.delete_key("desktop_icons")
	frappe.cache.delete_key("bootinfo")
	frappe.clear_cache()
