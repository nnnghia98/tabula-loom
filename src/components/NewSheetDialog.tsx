
interface NewSheetDialogProps {
  handleCreateSheet: (e: React.FormEvent) => void;
  newSheetDialogRef: React.RefObject<HTMLDialogElement | null>;
  newSheetName: string;
  setNewSheetName: React.Dispatch<React.SetStateAction<string>>;
}

export function NewSheetDialog({ handleCreateSheet, newSheetDialogRef, newSheetName, setNewSheetName }: NewSheetDialogProps) {
  return (
    <dialog ref={newSheetDialogRef} closedby="any" className="editorial-dialog">
      <form onSubmit={handleCreateSheet}>
        <div className="dialog-header">
          <h3 className="dialog-title">Add New Sheet</h3>
          <button
            type="button"
            className="btn btn-ghost btn-sm"
            onClick={() => newSheetDialogRef.current?.close()}
          >
            Close
          </button>
        </div>
        <div className="dialog-body">
          <div className="form-field">
            <label className="form-label">Sheet Name</label>
            <input
              type="text"
              className="input-control"
              placeholder="e.g. Sheet 2"
              value={newSheetName}
              onChange={(e) => setNewSheetName(e.target.value)}
              autoFocus
            />
          </div>
        </div>
        <div className="dialog-footer">
          <button
            type="button"
            className="btn"
            onClick={() => newSheetDialogRef.current?.close()}
          >
            Cancel
          </button>
          <button type="submit" className="btn btn-primary">
            Create Sheet
          </button>
        </div>
      </form>
    </dialog>
  );
}
