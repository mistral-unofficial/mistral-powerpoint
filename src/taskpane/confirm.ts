import type { ConfirmRequest } from "../tools";

/**
 * In-pane confirmation dialog. window.confirm() is blocked inside Office add-ins, so this renders
 * its own modal. Resolves true when the user confirms, false on Cancel, Escape or a backdrop click.
 */
export function confirmDialog({ title, message, confirmLabel }: ConfirmRequest): Promise<boolean> {
  const overlay = document.getElementById("confirm") as HTMLDivElement;
  const ok = document.getElementById("confirm-ok") as HTMLButtonElement;
  const cancel = document.getElementById("confirm-cancel") as HTMLButtonElement;
  document.getElementById("confirm-title")!.textContent = title;
  document.getElementById("confirm-message")!.textContent = message;
  ok.textContent = confirmLabel;

  const previousFocus = document.activeElement as HTMLElement | null;
  overlay.hidden = false;
  cancel.focus(); // the safe choice is the default

  return new Promise((resolve) => {
    const finish = (result: boolean) => {
      overlay.hidden = true;
      ok.removeEventListener("click", onOk);
      cancel.removeEventListener("click", onCancel);
      overlay.removeEventListener("click", onBackdrop);
      overlay.removeEventListener("keydown", onKey);
      previousFocus?.focus();
      resolve(result);
    };
    const onOk = () => finish(true);
    const onCancel = () => finish(false);
    const onBackdrop = (e: MouseEvent) => {
      if (e.target === overlay) finish(false);
    };
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") {
        e.preventDefault();
        finish(false);
      } else if (e.key === "Tab") {
        // Keep focus inside the dialog.
        e.preventDefault();
        (document.activeElement === ok ? cancel : ok).focus();
      }
    };
    ok.addEventListener("click", onOk);
    cancel.addEventListener("click", onCancel);
    overlay.addEventListener("click", onBackdrop);
    overlay.addEventListener("keydown", onKey);
  });
}
