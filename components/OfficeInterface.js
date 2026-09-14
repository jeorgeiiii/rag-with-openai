'use client';

import { useState } from 'react';

export default function OfficeInterface() {
  const [messages, setMessages] = useState([]);
  const [input, setInput] = useState('');
  const [loading, setLoading] = useState(false);

  async function handleSubmit(e) {
    e.preventDefault();
    if (!input.trim() || loading) return;

    const userMessage = { role: 'user', content: input };
    const history = messages.map(({ role, content }) => ({ role, content }));

    setMessages((prev) => [...prev, userMessage]);
    setInput('');
    setLoading(true);

    try {
      const csrfRes = await fetch('/api/csrf-token');
      const { csrfToken } = await csrfRes.json();

      const res = await fetch('/api/office-chat', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'X-CSRF-Token': csrfToken
        },
        body: JSON.stringify({ message: input, history })
      });

      const data = await res.json();

      if (res.ok) {
        setMessages((prev) => [
          ...prev,
          { role: 'assistant', content: data.response, actions: data.actions }
        ]);
      } else {
        setMessages((prev) => [
          ...prev,
          { role: 'assistant', content: `Error: ${data.error || 'Request failed'}`, error: true }
        ]);
      }
    } catch (error) {
      setMessages((prev) => [
        ...prev,
        { role: 'assistant', content: 'Error: Failed to connect to server', error: true }
      ]);
    } finally {
      setLoading(false);
    }
  }

  return (
    <div className="flex flex-col h-full">
      <div className="glass-card border-b p-4">
        <h3 className="font-semibold">Office Assistant (runs locally, Windows + Office required)</h3>
        <p className="text-sm opacity-80">
          Ask it to create/edit Excel workbooks, Word documents, PowerPoint presentations, Access databases, or
          Publisher publications. Files are kept in your Documents\ChatbotOffice folder.
        </p>
      </div>

      <div className="flex-1 overflow-y-auto p-4 space-y-4">
        {messages.length === 0 && (
          <div className="text-center opacity-90 mt-20 space-y-1">
            <p className="text-lg mb-2">Try one of these:</p>
            <p className="text-sm">&quot;Create a workbook called Budget with a Sheet1 and Sheet2&quot;</p>
            <p className="text-sm">&quot;Create a Word doc called Notes titled Meeting Notes&quot;</p>
            <p className="text-sm">&quot;Create a presentation called Pitch titled Our Product&quot;</p>
            <p className="text-sm">&quot;Create a database called Contacts with a People table&quot;</p>
            <p className="text-sm">&quot;Create a publication called Flyer with text Grand Opening&quot;</p>
          </div>
        )}

        {messages.map((msg, idx) => (
          <div key={idx} className={`flex ${msg.role === 'user' ? 'justify-end' : 'justify-start'}`}>
            <div
              className={`max-w-[80%] rounded-lg p-4 ${
                msg.role === 'user'
                  ? 'message-user'
                  : msg.error
                  ? 'bg-red-500/20 border border-red-500/40 text-white'
                  : 'message-assistant'
              }`}
            >
              <div className="whitespace-pre-wrap">{msg.content}</div>

              {msg.actions && msg.actions.length > 0 && (
                <div className="mt-3 pt-3 border-t border-white/20 space-y-1">
                  {msg.actions.map((a, i) => (
                    <div key={i} className="text-xs font-mono opacity-80">
                      {a.ok ? '✓' : '✗'} {a.tool}({Object.entries(a.args).map(([k, v]) => `${k}=${JSON.stringify(v)}`).join(', ')})
                      {!a.ok && a.result?.error && <span className="text-red-300"> — {a.result.error}</span>}
                    </div>
                  ))}
                </div>
              )}
            </div>
          </div>
        ))}

        {loading && (
          <div className="flex justify-start">
            <div className="glass-card rounded-lg p-4">
              <div className="flex space-x-2">
                <div className="w-2 h-2 bg-white/60 rounded-full animate-bounce"></div>
                <div className="w-2 h-2 bg-white/60 rounded-full animate-bounce" style={{ animationDelay: '0.1s' }}></div>
                <div className="w-2 h-2 bg-white/60 rounded-full animate-bounce" style={{ animationDelay: '0.2s' }}></div>
              </div>
            </div>
          </div>
        )}
      </div>

      <form onSubmit={handleSubmit} className="glass-card border-t p-4">
        <div className="flex gap-2">
          <input
            type="text"
            value={input}
            onChange={(e) => setInput(e.target.value)}
            placeholder="Tell it what to do in Excel..."
            className="flex-1 px-4 py-3 rounded-lg focus:outline-none focus:ring-2 focus:ring-purple-400"
            disabled={loading}
          />
          <button
            type="submit"
            disabled={loading || !input.trim()}
            className="glass-button px-6 py-3 rounded-lg disabled:opacity-50 disabled:cursor-not-allowed transition-colors font-medium"
          >
            Send
          </button>
        </div>
      </form>
    </div>
  );
}
