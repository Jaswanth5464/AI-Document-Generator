import { useState } from 'react';
import { collection, addDoc, serverTimestamp } from 'firebase/firestore';
import { db, auth } from '../firebase';

function CreateProjectModal({ onClose, onSuccess }) {
  const [projectName, setProjectName] = useState('');
  const [docType, setDocType] = useState('docx');
  const [loading, setLoading] = useState(false);

  const handleCreate = async (e) => {
    e.preventDefault();
    if (!projectName.trim()) return;

    setLoading(true);
    try {
      const userId = auth.currentUser.uid;
      const projectRef = await addDoc(collection(db, 'users', userId, 'projects'), {
        name: projectName,
        type: docType,
        topic: '',
        sections: [],
        createdAt: serverTimestamp(),
        status: 'draft'
      });

      onSuccess(projectRef.id, docType);
    } catch (error) {
      console.error('Error creating project:', error);
      alert('Failed to create project. Please try again.');
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal-content" onClick={(e) => e.stopPropagation()}>
        <div className="modal-header">
          <h2>📄 Create New Project</h2>
          <button className="btn-close" onClick={onClose}>×</button>
        </div>

        <form onSubmit={handleCreate}>
          <div className="form-group">
            <label>Project Name</label>
            <input
              type="text"
              placeholder="e.g., Market Analysis Report"
              value={projectName}
              onChange={(e) => setProjectName(e.target.value)}
              required
              autoFocus
            />
          </div>

          <div className="form-group">
            <label>Document Type</label>
            <div className="radio-group">
              <div className="radio-option">
                <input
                  type="radio"
                  id="docx"
                  name="docType"
                  value="docx"
                  checked={docType === 'docx'}
                  onChange={(e) => setDocType(e.target.value)}
                />
                <label htmlFor="docx" className="radio-label">
                  <span className="icon">📝</span>
                  <span className="text">Word Document</span>
                </label>
              </div>

              <div className="radio-option">
                <input
                  type="radio"
                  id="pptx"
                  name="docType"
                  value="pptx"
                  checked={docType === 'pptx'}
                  onChange={(e) => setDocType(e.target.value)}
                />
                <label htmlFor="pptx" className="radio-label">
                  <span className="icon">📊</span>
                  <span className="text">PowerPoint</span>
                </label>
              </div>
            </div>
          </div>

          <div className="modal-buttons">
            <button type="button" className="btn-secondary" onClick={onClose}>
              Cancel
            </button>
            <button type="submit" className="btn-primary" disabled={loading}>
              {loading ? '⏳ Creating...' : '✨ Create Project'}
            </button>
          </div>
        </form>
      </div>
    </div>
  );
}

export default CreateProjectModal;