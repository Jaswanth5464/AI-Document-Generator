import { useState, useEffect } from 'react';
import { signOut } from 'firebase/auth';
import { collection, query, getDocs, deleteDoc, doc } from 'firebase/firestore';
import { auth, db } from '../firebase';
import { useNavigate } from 'react-router-dom';
import CreateProjectModal from '../components/CreateProjectModal';
import ProjectCard from '../components/ProjectCard';

function Dashboard() {
  const [projects, setProjects] = useState([]);
  const [showModal, setShowModal] = useState(false);
  const [loading, setLoading] = useState(true);
  const navigate = useNavigate();

  useEffect(() => {
    loadProjects();
  }, []);

  const loadProjects = async () => {
    try {
      const userId = auth.currentUser.uid;
      const projectsRef = collection(db, 'users', userId, 'projects');
      const q = query(projectsRef);
      const snapshot = await getDocs(q);
      
      const projectsList = snapshot.docs.map(doc => ({
        id: doc.id,
        ...doc.data()
      }));
      
      setProjects(projectsList);
    } catch (error) {
      console.error('Error loading projects:', error);
    } finally {
      setLoading(false);
    }
  };

  const handleLogout = async () => {
    await signOut(auth);
    navigate('/login');
  };

  const handleCreateSuccess = (projectId, docType) => {
    setShowModal(false);
    if (docType === 'docx') {
      navigate(`/configure-word/${projectId}`);
    } else {
      navigate(`/configure-ppt/${projectId}`);
    }
  };

  const handleEditProject = (project) => {
    if (project.type === 'docx') {
      navigate(`/configure-word/${project.id}`);
    } else {
      navigate(`/configure-ppt/${project.id}`);
    }
  };

  const handleDeleteProject = async (projectId) => {
    try {
      const userId = auth.currentUser.uid;
      await deleteDoc(doc(db, 'users', userId, 'projects', projectId));
      loadProjects();
    } catch (error) {
      console.error('Error deleting project:', error);
      alert('Failed to delete project');
    }
  };

  return (
    <div className="dashboard">
      <div className="dashboard-header">
        <h1>📚 My Documents</h1>
        <button onClick={handleLogout} className="btn-logout">
          🚪 Logout
        </button>
      </div>

      <div className="dashboard-content">
        <button className="btn-create" onClick={() => setShowModal(true)}>
          ➕ Create New Project
        </button>

        {loading ? (
          <div style={{ textAlign: 'center', padding: '60px', fontSize: '20px' }}>
            ⏳ Loading projects...
          </div>
        ) : projects.length === 0 ? (
          <div className="empty-state">
            <div className="empty-state-icon">📭</div>
            <h2>No Projects Yet</h2>
            <p>Create your first AI-powered document to get started!</p>
          </div>
        ) : (
          <div className="projects-grid">
            {projects.map(project => (
              <ProjectCard
                key={project.id}
                project={project}
                onEdit={handleEditProject}
                onDelete={handleDeleteProject}
              />
            ))}
          </div>
        )}
      </div>

      {showModal && (
        <CreateProjectModal
          onClose={() => setShowModal(false)}
          onSuccess={handleCreateSuccess}
        />
      )}
    </div>
  );
}

export default Dashboard;