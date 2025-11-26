import API_URL from "../config";
import { useState, useEffect } from 'react';
import { useParams, useNavigate } from 'react-router-dom';
import { doc, getDoc, updateDoc } from 'firebase/firestore';
import { db, auth } from '../firebase';

function GenerateContent() {
  const { projectId } = useParams();
  const navigate = useNavigate();
  const [project, setProject] = useState(null);
  const [sections, setSections] = useState([]);
  const [currentIndex, setCurrentIndex] = useState(0);
  const [progress, setProgress] = useState(0);
  const [isGenerating, setIsGenerating] = useState(false);
  const [isComplete, setIsComplete] = useState(false);
  const [error, setError] = useState('');

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
        
        // Prepare sections with status
        const sectionsWithStatus = data.sections.map(s => ({
          ...s,
          status: 'pending' // pending, generating, completed, error
        }));
        setSections(sectionsWithStatus);
        
        // Auto-start generation after 1 second
        setTimeout(() => startGeneration(data, sectionsWithStatus), 1000);
      } else {
        setError('Project not found');
        setTimeout(() => navigate('/dashboard'), 2000);
      }
    } catch (error) {
      console.error('Error loading project:', error);
      setError('Failed to load project');
    }
  };

  const startGeneration = async (projectData, sectionsToGenerate) => {
    setIsGenerating(true);
    
    const updatedSections = [...sectionsToGenerate];
    
    for (let i = 0; i < updatedSections.length; i++) {
      setCurrentIndex(i);
      
      // Update status to generating
      updatedSections[i].status = 'generating';
      setSections([...updatedSections]);
      
      try {
        // Call backend to generate content
       const response = await fetch(`${API_URL}/api/generate-section`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            topic: projectData.topic,
            sectionTitle: updatedSections[i].title,
            docType: projectData.type
          })
        });
        
        if (!response.ok) throw new Error('Generation failed');
        
        const data = await response.json();
        
        // Update with generated content
        updatedSections[i].content = data.content;
        updatedSections[i].status = 'completed';
        setSections([...updatedSections]);
        
        // Update progress
        const newProgress = ((i + 1) / updatedSections.length) * 100;
        setProgress(newProgress);
        
      } catch (error) {
        console.error(`Error generating section ${i}:`, error);
        updatedSections[i].status = 'error';
        updatedSections[i].content = 'Failed to generate content. Please try again.';
        setSections([...updatedSections]);
      }
    }
    
    // Save to database
    try {
      const userId = auth.currentUser.uid;
      const projectRef = doc(db, 'users', userId, 'projects', projectId);
      
      await updateDoc(projectRef, {
        sections: updatedSections.map(s => ({
          id: s.id,
          title: s.title,
          content: s.content
        })),
        status: 'generated',
        lastModified: new Date()
      });
      
      setIsGenerating(false);
      setIsComplete(true);
      
      // Auto-navigate to refinement page after 2 seconds
      setTimeout(() => {
        navigate(`/refine/${projectId}`);
      }, 2000);
      
    } catch (error) {
      console.error('Error saving generated content:', error);
      setError('Failed to save content');
      setIsGenerating(false);
    }
  };

  if (error) {
    return (
      <div className="generate-container">
        <div className="generate-card">
          <div className="generate-icon">❌</div>
          <h2>Generation Failed</h2>
          <p>{error}</p>
          <button 
            className="btn-primary" 
            onClick={() => navigate('/dashboard')}
            style={{ marginTop: '20px' }}
          >
            Back to Dashboard
          </button>
        </div>
      </div>
    );
  }

  return (
    <div className="generate-container">
      <div className="generate-card">
        {!isComplete ? (
          <>
            <div className="generate-icon">🤖</div>
            <h2>Generating Your {project?.type === 'docx' ? 'Document' : 'Presentation'}</h2>
            <p>AI is creating amazing content for: <strong>{project?.topic}</strong></p>
            
            {/* Progress Bar */}
            <div className="progress-container">
              <div className="progress-bar">
                <div 
                  className="progress-fill" 
                  style={{ width: `${progress}%` }}
                />
              </div>
              <p className="progress-text">
                {Math.round(progress)}% Complete
              </p>
            </div>
            
            {/* Section Status List */}
            <div className="section-status-list">
              {sections.map((section, index) => (
                <div 
                  key={section.id} 
                  className={`section-status-item ${section.status}`}
                >
                  <span className="status-icon">
                    {section.status === 'completed' && '✅'}
                    {section.status === 'generating' && '⚡'}
                    {section.status === 'pending' && '⏳'}
                    {section.status === 'error' && '❌'}
                  </span>
                  <div className="status-text">
                    <div className="status-title">
                      {section.title}
                    </div>
                    <div className="status-subtitle">
                      {section.status === 'completed' && 'Content generated'}
                      {section.status === 'generating' && 'Generating now...'}
                      {section.status === 'pending' && 'Waiting...'}
                      {section.status === 'error' && 'Failed to generate'}
                    </div>
                  </div>
                </div>
              ))}
            </div>
          </>
        ) : (
          <>
            <div className="generate-icon success-animation">🎉</div>
            <h2>Generation Complete!</h2>
            <p>Your content is ready for refinement.</p>
            <p style={{ fontSize: '14px', color: '#666', marginTop: '10px' }}>
              Redirecting to editor...
            </p>
          </>
        )}
      </div>
    </div>
  );
}

export default GenerateContent;
