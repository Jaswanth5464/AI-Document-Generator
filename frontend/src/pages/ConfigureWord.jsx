import { useState, useEffect } from 'react';
import { useParams, useNavigate } from 'react-router-dom';
import { doc, getDoc, updateDoc } from 'firebase/firestore';
import { db, auth } from '../firebase';
import SectionEditor from '../components/SectionEditor';
import API_URL from '../config'; // ✅ NEW: Import API_URL

function ConfigureWord() {
  const { projectId } = useParams();
  const navigate = useNavigate();
  const [project, setProject] = useState(null);
  const [topic, setTopic] = useState('');
  const [sections, setSections] = useState([
    { id: 1, title: 'Introduction', content: '' },
    { id: 2, title: 'Main Content', content: '' },
    { id: 3, title: 'Conclusion', content: '' }
  ]);
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [suggesting, setSuggesting] = useState(false); // ✅ NEW: AI suggestion state

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
        setTopic(data.topic || '');
        if (data.sections && data.sections.length > 0) {
          setSections(data.sections);
        }
      } else {
        alert('Project not found!');
        navigate('/dashboard');
      }
    } catch (error) {
      console.error('Error loading project:', error);
      alert('Error loading project');
      navigate('/dashboard');
    } finally {
      setLoading(false);
    }
  };

  // ✅ NEW: AI Template Suggestion Function
  const handleAISuggest = async () => {
    if (!topic.trim()) {
      alert('⚠️ Please enter a topic first!');
      return;
    }

    setSuggesting(true);
    try {
      const response = await fetch(
        `${API_URL}/api/generate-template?topic=${encodeURIComponent(topic)}&doc_type=docx&num_sections=5`,
        { method: 'POST' }
      );

      if (!response.ok) throw new Error('Failed to generate template');

      const data = await response.json();
      
      // Replace current sections with AI suggestions
      setSections(data.sections);
      alert('✅ AI has suggested section titles! You can edit them if needed.');
      
    } catch (error) {
      console.error('Error generating template:', error);
      alert('❌ Failed to generate suggestions. Please try again.');
    } finally {
      setSuggesting(false);
    }
  };

  const handleGenerate = async () => {
    if (!topic.trim()) {
      alert('⚠️ Please enter a topic!');
      return;
    }

    const emptySections = sections.filter(s => !s.title.trim());
    if (emptySections.length > 0) {
      alert('⚠️ Please fill in all section titles!');
      return;
    }

    if (sections.length === 0) {
      alert('⚠️ Please add at least one section!');
      return;
    }

    setSaving(true);
    try {
      const userId = auth.currentUser.uid;
      const projectRef = doc(db, 'users', userId, 'projects', projectId);
      
      await updateDoc(projectRef, {
        topic,
        sections,
        status: 'configured',
        lastModified: new Date()
      });

      navigate(`/generate/${projectId}`);
    } catch (error) {
      console.error('Error saving:', error);
      alert('❌ Failed to save configuration. Please try again.');
    } finally {
      setSaving(false);
    }
  };

  if (loading) {
    return (
      <div style={{ 
        textAlign: 'center', 
        padding: '60px', 
        fontSize: '20px',
        minHeight: '100vh',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center'
      }}>
        ⏳ Loading project...
      </div>
    );
  }

  return (
    <div className="configure-container">
      <div className="configure-content">
        <div className="configure-card">
          <h2>📝 Configure Word Document</h2>
          <p className="subtitle">{project?.name}</p>

          <div className="form-group">
            <label>📋 Document Topic</label>
            <input
              type="text"
              placeholder="e.g., Market Analysis of Electric Vehicles in 2025"
              value={topic}
              onChange={(e) => setTopic(e.target.value)}
              style={{ fontSize: '16px' }}
            />
            <p className="password-hint">
              This will be the main topic for your document
            </p>
          </div>

          {/* ✅ NEW: AI Suggestion Button */}
          <div className="form-group">
            <button 
              className="btn-ai-suggest" 
              onClick={handleAISuggest}
              disabled={!topic.trim() || suggesting}
              type="button"
            >
              {suggesting ? '🤖 AI is thinking...' : '✨ AI Suggest Sections'}
            </button>
            <p className="password-hint">
              Let AI suggest section titles based on your topic
            </p>
          </div>

          <div className="form-group">
            <label>📑 Document Sections</label>
            <p className="password-hint" style={{ marginBottom: '15px' }}>
              Add, remove, or reorder sections as needed
            </p>
            <SectionEditor sections={sections} setSections={setSections} />
          </div>

          <div className="action-buttons">
            <button 
              className="btn-back" 
              onClick={() => navigate('/dashboard')}
              disabled={saving}
            >
              ← Back
            </button>
            <button 
              className="btn-generate" 
              onClick={handleGenerate}
              disabled={saving || !topic.trim() || sections.length === 0}
            >
              {saving ? '⏳ Saving...' : '🚀 Generate Content'}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

export default ConfigureWord;
