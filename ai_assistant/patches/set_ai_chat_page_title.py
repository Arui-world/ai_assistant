import frappe


def execute():
	if frappe.db.exists("Page", "ai-chat"):
		frappe.db.set_value(
			"Page",
			"ai-chat",
			{
				"title": "企业智能业务助手",
				"icon": "panel-top",
			},
			update_modified=False,
		)

		page = frappe.get_doc("Page", "ai-chat")
		page.set("roles", [])
		page.append("roles", {"role": "All"})
		page.flags.ignore_permissions = True
		page.save(ignore_permissions=True)
