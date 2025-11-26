import API_URL from "../config";
import { useState, useEffect } from 'react';
import { useParams, useNavigate } from 'react-router-dom';
import { doc, getDoc, updateDoc } from 'firebase/firestore';
import { db, auth } from '../firebase';

function RefineContent() {
  const { projectId } = useParams();
  const navigate = useNavigate();
  const [project, setProject] = useState(null);
  const [sections, setSections] = useState([]);
  const [loading, setLoading] = useState(true);
  const [exporting, setExporting] = useState(false);

  useEffect(() => {
    loadProject();
  }, [projectId]);

  const loadProject = async () => {
    try {
      const userId = auth.currentUser.uid;
      const projectRef = doc(db, 'users', userId, 'projects', projectId);
      const projectSnap = await getDoc(projectRef);
      
      if (projectSnap.exists()) {
        const data = projectSnap.data();
        setProject(data);
        setSections(data.sections || []);
      } else {
        alert('Project not found!');
        navigate('/dashboard');
      }
    } catch (error) {
      console.error('Error loading project:', error);
      alert('Failed to load project');
    } finally {
      setLoading(false);
    }
  };

  const handleRefine = async (sectionId, instruction) => {
    if (!instruction.trim()) {
      alert('⚠️ Please enter refinement instructions');
      return;
    }

    const sectionIndex = sections.findIndex(s => s.id === sectionId);
    if (sectionIndex === -1) return;

    // Show loading state
    const updatedSections = [...sections];
    updatedSections[sectionIndex].refining = true;
    setSections(updatedSections);

    try {
     const response = await fetch(`${API_URL}/api/refine-section`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          currentContent: sections[sectionIndex].content,
          instruction: instruction
        })
      });

      if (!response.ok) throw new Error('Refinement failed');

      const data = await response.json();
      
      // Update content
      updatedSections[sectionIndex].content = data.refinedContent;
      updatedSections[sectionIndex].refining = false;
      setSections(updatedSections);

      // Save to database
      await saveToDatabase(updatedSections);

      alert('✅ Content refined successfully!');
      
    } catch (error) {
      console.error('Error refining:', error);
      updatedSections[sectionIndex].refining = false;
      setSections(updatedSections);
      alert('❌ Failed to refine content. Please try again.');
    }
  };

  const handleFeedback = async (sectionId, feedbackType) => {
    const sectionIndex = sections.findIndex(s => s.id === sectionId);
    if (sectionIndex === -1) return;

    const updatedSections = [...sections];
    
    // Toggle feedback (click again to remove)
    if (updatedSections[sectionIndex].feedback === feedbackType) {
      updatedSections[sectionIndex].feedback = null;
    } else {
      updatedSections[sectionIndex].feedback = feedbackType;
    }
    
    setSections(updatedSections);
    await saveToDatabase(updatedSections);
  };

  const handleComment = async (sectionId, comment) => {
    const sectionIndex = sections.findIndex(s => s.id === sectionId);
    if (sectionIndex === -1) return;

    const updatedSections = [...sections];
    updatedSections[sectionIndex].comment = comment;
    setSections(updatedSections);
    
    await saveToDatabase(updatedSections);
  };

  const saveToDatabase = async (updatedSections) => {
    try {
      const userId = auth.currentUser.uid;
      const projectRef = doc(db, 'users', userId, 'projects', projectId);
      
      await updateDoc(projectRef, {
        sections: updatedSections.map(s => ({
          id: s.id,
          title: s.title,
          content: s.content,
          feedback: s.feedback || null,
          comment: s.comment || ''
        })),
        lastModified: new Date()
      });
    } catch (error) {
      console.error('Error saving to database:', error);
    }
  };

  const handleExport = async () => {
    setExporting(true);
    
    try {
      const response = await fetch(`${API_URL}/api/export-document`, {
        method: 'POST',
        headers: { 
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          topic: project.topic,
          sections: sections.map(s => ({
            id: s.id,
            title: s.title,
            content: s.content
          })),
          docType: project.type,
          theme: project.theme || "professional_blue"   // ✅ ADD THIS
        })
      });

      if (!response.ok) {
        const errorText = await response.text();
        console.error('Export error:', errorText);
        throw new Error('Export failed');
      }

      // Download file
      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${project.topic.replace(/\s+/g, '_')}.${project.type}`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);

      alert('✅ Document exported successfully!');
      
    } catch (error) {
      console.error('Error exporting:', error);
      alert('❌ Failed to export document. Please check console for details.');
    } finally {
      setExporting(false);
    }
  };

  if (loading) {
    return (
      <div style={{ 
        display: 'flex', 
        justifyContent: 'center', 
        alignItems: 'center', 
        height: '100vh',
        fontSize: '24px'
      }}>
        ⏳ Loading...
      </div>
    );
  }

  return (
    <div className="refine-container">
      {/* Header */}
      <div className="refine-header">
        <h1>
          <span>✏️</span> Refine Your Content
        </h1>
        <div className="header-actions">
          <button 
            className="btn-secondary" 
            onClick={() => navigate('/dashboard')}
          >
            ← Back to Dashboard
          </button>
          <button 
            className="btn-export" 
            onClick={handleExport}
            disabled={exporting}
          >
            {exporting ? (
              <>
                <span className="spinner"></span>
                Exporting...
              </>
            ) : (
              <>
                <span>📥</span>
                Export {project?.type === 'docx' ? 'Word' : 'PowerPoint'}
              </>
            )}
          </button>
        </div>
      </div>

      {/* Content */}
      <div className="refine-content">
        <div style={{ marginBottom: '30px', textAlign: 'center' }}>
          <h2 style={{ color: '#333', marginBottom: '10px' }}>
            {project?.topic}
          </h2>
          <p style={{ color: '#666' }}>
            Review and refine each {project?.type === 'docx' ? 'section' : 'slide'} below
          </p>
        </div>

        {sections.map((section, index) => (
          <SectionCard
            key={section.id}
            section={section}
            index={index}
            docType={project?.type}
            onRefine={handleRefine}
            onFeedback={handleFeedback}
            onComment={handleComment}
          />
        ))}
      </div>
    </div>
  );
}

// ========== SECTION CARD COMPONENT ==========
function SectionCard({ section, index, docType, onRefine, onFeedback, onComment }) {
  const [refineInput, setRefineInput] = useState('');
  const [commentInput, setCommentInput] = useState(section.comment || '');
  const [showComment, setShowComment] = useState(false);

  const handleRefineClick = () => {
    if (!refineInput.trim()) {
      alert('⚠️ Please enter refinement instructions');
      return;
    }
    onRefine(section.id, refineInput);
    setRefineInput('');
  };

  const handleCommentSave = () => {
    onComment(section.id, commentInput);
    alert('💬 Comment saved!');
  };

  return (
    <div className="section-card">
      <div className="section-card-header">
        <span className="section-number-badge">
          {docType === 'docx' ? 'Section' : 'Slide'} {index + 1}
        </span>
      </div>

      <h3>{section.title}</h3>

      <div className="section-content">
        {section.content}
      </div>

      {/* Refinement Controls */}
      <div className="refinement-controls">
        <input
          type="text"
          className="refine-input"
          placeholder="e.g., 'Make it shorter', 'Add more details', 'Make it more formal'"
          value={refineInput}
          onChange={(e) => setRefineInput(e.target.value)}
          onKeyPress={(e) => {
            if (e.key === 'Enter') {
              handleRefineClick();
            }
          }}
          disabled={section.refining}
        />
        <button 
          className="btn-refine" 
          onClick={handleRefineClick}
          disabled={section.refining}
        >
          {section.refining ? (
            <>
              <span className="spinner"></span>
              Refining...
            </>
          ) : (
            '✨ Refine with AI'
          )}
        </button>
      </div>

      {/* Feedback Actions */}
      <div className="feedback-actions">
        <button 
          className={`btn-feedback ${section.feedback === 'like' ? 'liked' : ''}`}
          onClick={() => onFeedback(section.id, 'like')}
        >
          👍 {section.feedback === 'like' ? 'Liked' : 'Like'}
        </button>
        <button 
          className={`btn-feedback ${section.feedback === 'dislike' ? 'disliked' : ''}`}
          onClick={() => onFeedback(section.id, 'dislike')}
        >
          👎 {section.feedback === 'dislike' ? 'Disliked' : 'Dislike'}
        </button>
        <button 
          className="btn-feedback"
          onClick={() => setShowComment(!showComment)}
        >
          💬 {showComment ? 'Hide Comment' : 'Add Comment'}
        </button>
      </div>

      {/* Comment Section */}
      {showComment && (
        <div className="comment-section">
          <textarea
            className="comment-textarea"
            placeholder="Add your notes or comments about this section..."
            value={commentInput}
            onChange={(e) => setCommentInput(e.target.value)}
          />
          <button 
            className="btn-save-comment"
            onClick={handleCommentSave}
          >
            💾 Save Comment
          </button>
        </div>
      )}
    </div>
  );
}

export default RefineContent;
