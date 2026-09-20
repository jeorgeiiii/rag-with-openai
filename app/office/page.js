import OfficeInterface from '../../components/OfficeInterface';

export default function OfficePage() {
  // The assistant drives desktop Office through Windows COM, which cannot run on Vercel (Linux).
  const unsupported = process.platform !== 'win32';

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

          {unsupported ? (
            <div className="glass-card rounded-lg p-8 text-center">
              <h2 className="text-2xl font-semibold mb-3">⚠️ Office Assistant is not available here</h2>
              <p className="mb-3">
                This app is deployed on Vercel (a Linux server), which has no Microsoft Office
                and cannot run Windows COM automation.
              </p>
              <p className="mb-3">
                This is a personal project built to make developers&apos; work easier, so the
                Office Assistant is mostly not a public feature. It&apos;s meant for people
                running the project on their own machine.
              </p>
              <p>
                To use this feature, clone the project from GitHub and run it locally on a
                Windows PC with Microsoft Office installed (<code>npm install</code> then{' '}
                <code>npm run dev</code>).
              </p>
            </div>
          ) : (
            <div className="glass-card rounded-lg flex flex-col overflow-hidden h-[700px]">
              <OfficeInterface />
            </div>
          )}
        </div>
      </div>
    </main>
  );
}
