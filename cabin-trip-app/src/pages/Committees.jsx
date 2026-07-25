import { useState } from 'react';
import { useIdentity } from '../identity';
import { usePoll } from '../usePoll';
import { api } from '../api';
import { nameOf } from '../utils';

export default function Committees({ people }) {
  const identity = useIdentity();
  const committees = usePoll('/committees', 8000);
  const peopleList = people.data || [];
  const list = committees.data || [];
  const [newName, setNewName] = useState('');
  const [newEmoji, setNewEmoji] = useState('🗂️');

  async function addCommittee() {
    if (!newName.trim()) return;
    await api.post('/committees', { name: newName.trim(), emoji: newEmoji.trim() || '🗂️' });
    setNewName(''); setNewEmoji('🗂️');
    committees.reload();
  }
  async function removeCommittee(c) {
    if (!confirm(`Delete "${c.name}" committee? This removes its tasks and notes too.`)) return;
    await api.del(`/committees/${c.id}`);
    committees.reload();
  }
  async function join(c) {
    const inIt = c.members.includes(identity.personId);
    await api.post(`/committees/${c.id}/${inIt ? 'leave' : 'join'}`, {});
    committees.reload();
  }
  async function toggleTask(t) {
    await api.patch(`/committee-tasks/${t.id}`, { done: !t.done });
    committees.reload();
  }
  async function addTask(c) {
    const task = prompt('New task:');
    if (!task) return;
    await api.post(`/committees/${c.id}/tasks`, { task });
    committees.reload();
  }
  async function addNote(c) {
    const text = prompt('Add a note:');
    if (!text) return;
    await api.post(`/committees/${c.id}/notes`, { text });
    committees.reload();
  }

  return (
    <div>
      {list.map(c => (
        <div key={c.id} className="card">
          <div className="spread">
            <h3>{c.emoji} {c.name}</h3>
            <span className="row">
              <button className={`btn small ${c.members.includes(identity.personId) ? '' : 'ghost'}`} onClick={() => join(c)}>
                {c.members.includes(identity.personId) ? 'Leave' : 'Join'}
              </button>
              <button className="btn small danger" onClick={() => removeCommittee(c)}>×</button>
            </span>
          </div>
          <div className="small-muted">Members: {c.members.map(id => nameOf(peopleList, id)).join(', ') || 'none yet'}</div>

          <div className="section-title"><strong>Tasks</strong></div>
          {c.tasks.map(t => (
            <div key={t.id} className={`checklist-item ${t.done ? 'done' : ''}`}>
              <input type="checkbox" checked={!!t.done} onChange={() => toggleTask(t)} />
              <span style={{ flex: 1 }}>{t.task} {t.assigneeId ? `— ${nameOf(peopleList, t.assigneeId)}` : ''} {t.dueDate ? `(due ${t.dueDate})` : ''}</span>
            </div>
          ))}
          <button className="btn small ghost" onClick={() => addTask(c)}>+ Task</button>

          <div className="section-title"><strong>Notes</strong></div>
          <div className="notes-thread">
            {c.notes.map(n => (
              <div key={n.id} className="note-bubble">
                <div className="meta">{nameOf(peopleList, n.personId) || 'Someone'} · {new Date(n.ts).toLocaleDateString()}</div>
                {n.text}
              </div>
            ))}
            {c.notes.length === 0 && <p className="small-muted">No notes yet.</p>}
          </div>
          <button className="btn small ghost" onClick={() => addNote(c)}>+ Note</button>
        </div>
      ))}

      <div className="card">
        <h3>+ New Committee</h3>
        <div className="row">
          <input type="text" placeholder="Emoji" style={{ width: 60 }} value={newEmoji} onChange={e => setNewEmoji(e.target.value)} />
          <input type="text" placeholder="Committee name" value={newName} onChange={e => setNewName(e.target.value)} />
        </div>
        <button className="btn small" style={{ marginTop: 8 }} onClick={addCommittee}>Add committee</button>
      </div>
    </div>
  );
}
