(function () {
	const BUTTON_ID = "ai-assistant-navbar-link";
	const AI_CHAT_URL = "/desk/ai-chat";
	const BUTTON_LABEL = "ai小助手";

	function goToAiChat(event) {
		event.preventDefault();
		if (window.frappe && typeof frappe.set_route === "function") {
			Promise.resolve(frappe.set_route("ai-chat")).catch(function () {
				window.location.href = AI_CHAT_URL;
			});
		} else {
			window.location.href = AI_CHAT_URL;
		}
	}

	function injectButton() {
		if (document.getElementById(BUTTON_ID)) {
			return true;
		}

		var $bellBtn = $(".desktop-notifications .btn-reset.nav-link[data-toggle='dropdown']");
		if (!$bellBtn.length) {
			return false;
		}

		var $link = $("<a>", {
			id: BUTTON_ID,
			href: AI_CHAT_URL,
			"class": "btn btn-secondary btn-sm"
		}).css({
			whiteSpace: "nowrap",
			lineHeight: "1.4"
		}).on("click", function (e) {
			e.preventDefault();
			goToAiChat(e);
		}).text(BUTTON_LABEL);

		$bellBtn.parent().parent().before($link);
		return true;
	}

	function scheduleInjection(retryCount) {
		if (injectButton() || retryCount <= 0) {
			return;
		}
		window.setTimeout(function () {
			scheduleInjection(retryCount - 1);
		}, 300);
	}

	$(document).ready(function () {
		scheduleInjection(10);
	});

	if (window.frappe && frappe.router) {
		frappe.router.on("change", function () {
			scheduleInjection(10);
		});
	}
})();
