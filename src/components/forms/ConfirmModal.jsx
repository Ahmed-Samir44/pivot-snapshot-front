import Modal from "./Modal";

// A styled stand-in for window.confirm() — some environments/policies block or discourage native
// browser confirm() popups (flagged live, 2026-09-23), and it never matched this app's own look
// anyway. Reuses Modal.jsx for the overlay/centering/Escape-to-close behavior.
export default function ConfirmModal({ title, message, confirmLabel = "Confirm", danger = false, onConfirm, onCancel }) {
  return (
    <Modal title={title} onClose={onCancel}>
      <p className="mb-5 max-w-sm text-sm text-ink">{message}</p>
      <div className="flex justify-end gap-3">
        <button type="button" onClick={onCancel} className="btn-secondary">
          Cancel
        </button>
        <button
          type="button"
          onClick={onConfirm}
          className={danger ? "rounded-xl bg-red-600 px-4 py-2 text-sm font-semibold text-white hover:bg-red-700" : "btn-primary"}
        >
          {confirmLabel}
        </button>
      </div>
    </Modal>
  );
}
