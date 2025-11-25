function SectionEditor({ sections, setSections }) {
  const addSection = () => {
    const newSection = {
      id: Date.now(),
      title: '',
      content: ''
    };
    setSections([...sections, newSection]);
  };

  const removeSection = (id) => {
    setSections(sections.filter(s => s.id !== id));
  };

  const updateSection = (id, title) => {
    setSections(sections.map(s => 
      s.id === id ? { ...s, title } : s
    ));
  };

  const moveSection = (index, direction) => {
    const newSections = [...sections];
    const newIndex = index + direction;
    
    if (newIndex < 0 || newIndex >= sections.length) return;
    
    [newSections[index], newSections[newIndex]] = [newSections[newIndex], newSections[index]];
    setSections(newSections);
  };

  return (
    <div className="section-list">
      {sections.map((section, index) => (
        <div key={section.id} className="section-item">
          <div className="section-header">
            <span className="drag-handle">⋮⋮</span>
            <span className="section-number">{index + 1}</span>
            <input
              type="text"
              className="section-input"
              placeholder="e.g., Introduction, Market Analysis..."
              value={section.title}
              onChange={(e) => updateSection(section.id, e.target.value)}
            />
            <button
              type="button"
              className="btn-icon"
              onClick={() => moveSection(index, -1)}
              disabled={index === 0}
              title="Move up"
            >
              ⬆️
            </button>
            <button
              type="button"
              className="btn-icon"
              onClick={() => moveSection(index, 1)}
              disabled={index === sections.length - 1}
              title="Move down"
            >
              ⬇️
            </button>
            <button
              type="button"
              className="btn-icon btn-remove"
              onClick={() => removeSection(section.id)}
              title="Remove section"
            >
              🗑️
            </button>
          </div>
        </div>
      ))}

      <button type="button" className="btn-add-section" onClick={addSection}>
        ➕ Add Section
      </button>
    </div>
  );
}

export default SectionEditor;