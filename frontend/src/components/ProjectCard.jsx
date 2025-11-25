function ProjectCard({ project, onEdit, onDelete }) {
  const formatDate = (timestamp) => {
    if (!timestamp) return 'Just now';
    const date = timestamp.toDate();
    return date.toLocaleDateString('en-US', { 
      month: 'short', 
      day: 'numeric',
      year: 'numeric'
    });
  };

  return (
    <div className="project-card">
      <div className="project-card-header">
        <span className="project-icon">
          {project.type === 'docx' ? '📝' : '📊'}
        </span>
        <span className={`project-type-badge badge-${project.type}`}>
          {project.type === 'docx' ? 'WORD' : 'PPT'}
        </span>
      </div>

      <h3>{project.name}</h3>
      <p>{project.topic || 'No topic set yet'}</p>
      <p className="project-date">Created: {formatDate(project.createdAt)}</p>

      <div className="project-actions">
        <button className="btn-small btn-edit" onClick={() => onEdit(project)}>
          ✏️ Edit
        </button>
        <button 
          className="btn-small btn-delete" 
          onClick={(e) => {
            e.stopPropagation();
            if (window.confirm('Delete this project?')) {
              onDelete(project.id);
            }
          }}
        >
          🗑️
        </button>
      </div>
    </div>
  );
}

export default ProjectCard;