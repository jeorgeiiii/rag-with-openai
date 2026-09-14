import OfficeInterface from '../../components/OfficeInterface';

export default function OfficePage() {
  return (
    <main className="min-h-screen w-full">
      <a
        href="/"
        className="nav-btn top-5 left-5 z-50 px-4 py-2 rounded-lg text-sm font-medium hidden md:inline-block"
      >
        Back to RAG Chatbot
      </a>

      <div className="container mx-auto px-4 py-8">
        <div className="max-w-5xl mx-auto w-full">
          <div className="text-center mb-6">
            <h1 className="text-4xl font-bold mb-2">Office Assistant</h1>
            <p className="opacity-90">Chat to create, edit, and read Excel, Word, and PowerPoint files via COM automation</p>
          </div>

          <div className="glass-card rounded-lg flex flex-col overflow-hidden h-[700px]">
            <OfficeInterface />
          </div>
        </div>
      </div>
    </main>
  );
}
