window.msalRedirectBridge.broadcastResponseToMainFrame().catch(() => {
  document.body.textContent = "Microsoft sign-in could not be completed. You can close this window and try again.";
});
