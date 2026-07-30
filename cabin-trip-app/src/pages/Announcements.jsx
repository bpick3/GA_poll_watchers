import { useState } from 'react';
import { useIdentity } from '../identity';
import { usePoll } from '../usePoll';
import { api } from '../api';
import { nameOf, shareText } from '../utils';

const TEMPLATES = [
  { label: '💵 Payment reminder', text: "Reminder: don't forget to send your cabin trip payment! " },
  { label: '📅 Schedule change', text: "Heads up — schedule change: " },
  { label: '🚗 Travel update', text: "Travel update: " },
  { label: '🎉 General', text: '' },
];

export default function Announcements({ people, settings }) {
  const identity = useIdentity();
  const announcements = usePoll('/announcements', 8000);
  const peopleList = people.data || [];
  const tripName = (settings?.data || {}).tripName || 'Cabin trip';
  const [text, setText] = useState('');
  const [status, setStatus] = useState('');

  async function post() {
    if (!text.trim()) return;
    await api.post('/announcements', { text: text.trim() });
    setText('');
    announcements.reload();
  }

  async function remove(id) {
    await api.del(`/announcements/${id}`);
    announcements.reload();
  }

  async function share(msgText) {
    const result = await shareText(msgText, tripName);
    if (result === 'copied') {
      setStatus('Copied to clipboard!');
      setTimeout(() => setStatus(''), 2500);
    } else if (result === 'unsupported') {
      setStatus("Couldn't share automatically — copy the text manually.");
      setTimeout(() => setStatus(''), 3000);
    }
  }

  const list = announcements.data || [];

  return (
    <div>
      <div className="card">
        <h3>📢 New Announcement</h3>
        <p className="small-muted">Post it here for the group to see in-app, and/or tap Share to send it as a text through your phone's own Messages app.</p>
        <div className="chip-row">
          {TEMPLATES.map(t => (
            <span key={t.label} className="chip" style={{ cursor: 'pointer' }} onClick={() => setText(t.text)}>{t.label}</span>
          ))}
        </div>
        <div className="field" style={{ marginTop: 8 }}>
          <textarea value={text} onChange={e => setText(e.target.value)} placeholder="Write your message…" />
        </div>
        <div className="row">
          <button className="btn small" onClick={post}>Post</button>
          <button className="btn small secondary" disabled={!text.trim()} onClick={() => share(text.trim())}>📤 Share as text</button>
        </div>
        {status && <div className="small-muted" style={{ marginTop: 6 }}>{status}</div>}
      </div>

      <div className="card">
        <h3>Recent Announcements</h3>
        {list.length === 0 && <p className="small-muted">No announcements yet.</p>}
        {list.map(a => (
          <div key={a.id} className="note-bubble" style={{ marginBottom: 8 }}>
            <div className="meta">{nameOf(peopleList, a.personId) || 'Someone'} · {new Date(a.ts).toLocaleString()}</div>
            <div>{a.text}</div>
            <div className="row" style={{ marginTop: 6 }}>
              <button className="btn small ghost" onClick={() => share(a.text)}>📤 Share as text</button>
              <button className="btn small danger" onClick={() => remove(a.id)}>Delete</button>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}
