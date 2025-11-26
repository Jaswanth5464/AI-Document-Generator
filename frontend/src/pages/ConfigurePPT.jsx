import { useState, useEffect } from 'react';
import { useParams, useNavigate } from 'react-router-dom';
import { doc, getDoc, updateDoc } from 'firebase/firestore';
import { db, auth } from '../firebase';
import SectionEditor from '../components/SectionEditor';
import API_URL from '../config'; // ✅ NEW: Import API_URL

function ConfigurePPT() {
  const { projectId } = useParams();
  const navigate = useNavigate();
  const [project, setProject] = useState(null);
  const [topic, setTopic] = useState('');
  const [sections, setSections] = useState([
    { id: 1, title: 'Title Slide', content: '' },
    { id: 2, title: 'Overview', content: '' },
    { id: 3, title: 'Key Points', content: '' },
    { id: 4, title: 'Conclusion', content: '' }
  ]);
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [suggesting, setSuggesting] = useState(false); // ✅ NEW: AI suggestion state
  const [selectedTheme, setSelectedTheme] = useState('professional_blue'); // ADD THIS LINE
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
      alert('⚠️ Please enter a presentation topic first!');
      return;
    }

    setSuggesting(true);
    try {
      const response = await fetch(
        `${API_URL}/api/generate-template?topic=${encodeURIComponent(topic)}&doc_type=pptx&num_sections=5`,
        { method: 'POST' }
      );

      if (!response.ok) throw new Error('Failed to generate template');

      const data = await response.json();
      
      // Replace current sections with AI suggestions
      setSections(data.sections);
      alert('✅ AI has suggested slide titles! You can edit them if needed.');
      
    } catch (error) {
      console.error('Error generating template:', error);
      alert('❌ Failed to generate suggestions. Please try again.');
    } finally {
      setSuggesting(false);
    }
  };

  const handleGenerate = async () => {
  if (!topic.trim()) {
    alert('⚠️ Please enter a presentation topic!');
    return;
  }

  const emptySlides = sections.filter(s => !s.title.trim());
  if (emptySlides.length > 0) {
    alert('⚠️ Please fill in all slide titles!');
    return;
  }

  if (sections.length === 0) {
    alert('⚠️ Please add at least one slide!');
    return;
  }

  setSaving(true);
  try {
    const userId = auth.currentUser.uid;
    const projectRef = doc(db, 'users', userId, 'projects', projectId);
    
    await updateDoc(projectRef, {
      topic,
      sections,
      theme: selectedTheme, // ADD THIS LINE
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
          <h2>📊 Configure PowerPoint Presentation</h2>
          <p className="subtitle">{project?.name}</p>

          <div className="form-group">
            <label>📋 Presentation Topic</label>
            <input
              type="text"
              placeholder="e.g., Future of Artificial Intelligence in Business"
              value={topic}
              onChange={(e) => setTopic(e.target.value)}
              style={{ fontSize: '16px' }}
            />
            <p className="password-hint">
              This will be the main topic for your presentation
            </p>
          </div>

          {/* ✅ AI Suggestion Button */}
          <div className="form-group">
            <button
              className="btn-ai-suggest"
              onClick={handleAISuggest}
              disabled={!topic.trim() || suggesting}
              type="button"
            >
              {suggesting ? '🤖 AI is thinking...' : '✨ AI Suggest Slides'}
            </button>
            <p className="password-hint">
              Let AI suggest slide titles based on your topic
            </p>
          </div>

          {/* 🎨 =============== THEME SELECTOR =============== */}
          <div className="form-group">
            <label>🎨 Presentation Theme</label>
            <select
              value={selectedTheme}
              onChange={(e) => setSelectedTheme(e.target.value)}
              className="theme-selector"
            >
              <option value="professional_blue">💙 Professional Blue</option>
              <option value="modern_dark">🖤 Modern Dark</option>
              <option value="vibrant_orange">🧡 Vibrant Orange</option>
              <option value="nature_green">💚 Nature Green</option>
              <option value="elegant_purple">💜 Elegant Purple</option>
            </select>

            <p className="password-hint">
              Choose a color theme for your presentation
            </p>

            <div className={`theme-preview theme-${selectedTheme}`}>
              <div className="preview-title">Preview</div>
              <div className="preview-content">
                <div className="preview-bullet">• First point</div>
                <div className="preview-bullet">• Second point</div>
              </div>
            </div>
          </div>
          {/* 🎨 =============== END THEME SELECTOR =============== */}

          <div className="form-group">
            <label>🎬 Slide Titles</label>
            <p className="password-hint" style={{ marginBottom: '15px' }}>
              Add, remove, or reorder slides as needed. Each slide will get AI-generated content.
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

export default ConfigurePPT;
