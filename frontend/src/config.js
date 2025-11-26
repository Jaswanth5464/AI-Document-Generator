const API_URL = import.meta.env.PROD
  ? "https://ai-doc-generator-backend.onrender.com"
  : "http://localhost:8000";

export default API_URL;
