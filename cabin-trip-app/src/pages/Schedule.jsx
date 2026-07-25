import { useState } from 'react';
import { useIdentity } from '../identity';
import { usePoll } from '../usePoll';
import { api } from '../api';
import { nameOf, initials } from '../utils';

const ALT_THEMES_FALLBACK = ['Y2K Night 💿', 'Cabin Casino 🎰', 'Murder Mystery Dinner 🔪', 'White Lies Party 🤥', 'Decades Night 🕺'];

export default function Schedule({ people }) {
  const identity = useIdentity();
  const days = usePoll('/days', 8000);
  const settings = usePoll('/settings', 20000);
  const [activeDay, setActiveDay] = useState(0);
  const [showAddBlock, setShowAddBlock] = useState(false);

  const peopleList = people.data || [];
  const dayList = days.data || [];
  const day = dayList[activeDay];

  const altThemes = (() => {
    try { return JSON.parse((settings.data || {}).altThemes || '[]'); } catch { return ALT_THEMES_FALLBACK; }
  })();

  async function rsvp(blockId, status) {
    await api.post(`/blocks/${blockId}/rsvp`, { status });
    days.reload();
  }

  async function changeTheme(dayId, theme) {
    await api.patch(`/days/${dayId}`, { theme });
    days.reload();
  }

  async function vote(nomId) {
    try {
      await api.post(`/movie-nominations/${nomId}/vote`, {});
      days.reload();
    } catch (e) { alert(e.message); }
  }

  async function nominate(dayId) {
    const title = prompt('Nominate a movie:');
    if (!title) return;
    await api.post('/movie-nominations', { dayId, title, nominatedBy: identity.personId });
    days.reload();
  }

  async function toggleHotTub(slotId) {
    try {
      await api.post(`/hottub-slots/${slotId}/toggle`, {});
      days.reload();
    } catch (e) { alert(e.message); }
  }

  async function deleteBlock(id) {
    if (!confirm('Delete this block?')) return;
    await api.del(`/blocks/${id}`);
    days.reload();
  }

  if (!day) return <p className="small-muted">Loading schedule…</p>;

  return (
    <div>
      <div className="day-tabs">
        {dayList.map((d, i) => (
          <button key={d.id} className={i === activeDay ? 'active' : ''} onClick={() => setActiveDay(i)}>{d.label}</button>
        ))}
      </div>

      <div className="card">
        <div className="spread">
          <h3>{day.theme}</h3>
        </div>
        <div className="field">
          <label>Pick from theme library</label>
          <select value={day.theme} onChange={e => changeTheme(day.id, e.target.value)}>
            <option value={day.theme}>{day.theme} (current)</option>
            {altThemes.filter(t => t !== day.theme).map(t => <option key={t} value={t}>{t}</option>)}
          </select>
        </div>
        <CustomThemeInput dayId={day.id} onSave={changeTheme} />
      </div>

      {(day.blocks || []).map(block => (
        <BlockCard key={block.id} block={block} people={peopleList} identity={identity} onRsvp={rsvp} onDelete={deleteBlock} reload={days.reload} />
      ))}

      {day.movieNominations && (
        <MovieNight day={day} onVote={vote} onNominate={nominate} identity={identity} people={peopleList} />
      )}

      {day.hottubSlots && day.hottubSlots.length > 0 && (
        <HotTub day={day} onToggle={toggleHotTub} identity={identity} people={peopleList} />
      )}

      <button className="btn secondary" onClick={() => setShowAddBlock(true)}>+ Add a Block</button>
      {showAddBlock && <AddBlockForm dayId={day.id} people={peopleList} onClose={() => { setShowAddBlock(false); days.reload(); }} />}
    </div>
  );
}

function BlockCard({ block, people, identity, onRsvp, onDelete, reload }) {
  const [editing, setEditing] = useState(false);
  const inList = (block.rsvps || []).filter(r => r.status === 'in');
  const myRsvp = (block.rsvps || []).find(r => r.personId === identity.personId);
  const owner = block.ownerId ? nameOf(people, block.ownerId) : (block.ownerText || null);
  const noOwner = !block.ownerId && !block.ownerText;

  return (
    <div className={`card block-card ${noOwner ? 'no-owner' : ''}`}>
      <div className="spread">
        <div>
          <strong>{block.time ? `${block.time} — ` : ''}{block.title}</strong>
          {block.location && <div className="small-muted">📍 {block.location}</div>}
        </div>
        <button className="btn small ghost" onClick={() => setEditing(v => !v)}>⋯</button>
      </div>
      <div className="small-muted">Owner: {owner || <span className="needs-owner">unclaimed</span>}</div>
      {block.notes && <div className="small-muted" style={{ marginTop: 4 }}>{block.notes}</div>}
      <div className="row" style={{ marginTop: 8 }}>
        <div className="avatar-row">
          {inList.map(r => <div key={r.personId} className="avatar" title={nameOf(people, r.personId)}>{initials(nameOf(people, r.personId))}</div>)}
        </div>
        <div style={{ marginLeft: 'auto' }} className="row">
          <button className={`btn small ${myRsvp?.status === 'in' ? '' : 'ghost'}`} onClick={() => onRsvp(block.id, 'in')}>I'm in</button>
          <button className={`btn small ${myRsvp?.status === 'out' ? 'danger' : 'ghost'}`} onClick={() => onRsvp(block.id, 'out')}>I'm out</button>
        </div>
      </div>
      {editing && (
        <EditBlock block={block} people={people} onDone={() => { setEditing(false); reload(); }} onDelete={() => onDelete(block.id)} />
      )}
    </div>
  );
}

function EditBlock({ block, people, onDone, onDelete }) {
  const [ownerId, setOwnerId] = useState(block.ownerId || '');
  const [ownerText, setOwnerText] = useState(block.ownerText || '');
  const [notes, setNotes] = useState(block.notes || '');

  async function save() {
    await api.patch(`/blocks/${block.id}`, { ownerId: ownerId || null, ownerText: ownerId ? '' : ownerText, notes });
    onDone();
  }

  return (
    <div style={{ marginTop: 10, borderTop: '1px solid var(--cream-2)', paddingTop: 10 }}>
      <div className="field">
        <label>Assign owner (person)</label>
        <select value={ownerId} onChange={e => setOwnerId(e.target.value)}>
          <option value="">— none —</option>
          {people.map(p => <option key={p.id} value={p.id}>{p.name}</option>)}
        </select>
      </div>
      {!ownerId && (
        <div className="field">
          <label>Or free-text owner</label>
          <input type="text" value={ownerText} onChange={e => setOwnerText(e.target.value)} placeholder="e.g. Everyone / TBD" />
        </div>
      )}
      <div className="field">
        <label>Notes</label>
        <textarea value={notes} onChange={e => setNotes(e.target.value)} />
      </div>
      <div className="row">
        <button className="btn small" onClick={save}>Save</button>
        <button className="btn small danger" onClick={onDelete}>Delete Block</button>
      </div>
    </div>
  );
}

function CustomThemeInput({ dayId, onSave }) {
  const [text, setText] = useState('');
  async function save() {
    if (!text.trim()) return;
    await onSave(dayId, text.trim());
    setText('');
  }
  return (
    <div className="field">
      <label>Or type a custom theme</label>
      <div className="row">
        <input type="text" placeholder="e.g. Board Game Bonanza 🎲" value={text} onChange={e => setText(e.target.value)} />
        <button className="btn small" onClick={save}>Set</button>
      </div>
    </div>
  );
}

function AddBlockForm({ dayId, people, onClose }) {
  const [title, setTitle] = useState('');
  const [time, setTime] = useState('');
  const [location, setLocation] = useState('');

  async function save() {
    if (!title.trim()) return;
    await api.post('/blocks', { dayId, title, time, location });
    onClose();
  }

  return (
    <div className="card">
      <h3>New Block</h3>
      <div className="field"><label>Title</label><input type="text" value={title} onChange={e => setTitle(e.target.value)} /></div>
      <div className="field"><label>Time</label><input type="text" placeholder="e.g. 4:00 PM" value={time} onChange={e => setTime(e.target.value)} /></div>
      <div className="field"><label>Location</label><input type="text" value={location} onChange={e => setLocation(e.target.value)} /></div>
      <div className="row">
        <button className="btn small" onClick={save}>Add</button>
        <button className="btn small ghost" onClick={onClose}>Cancel</button>
      </div>
    </div>
  );
}

function MovieNight({ day, onVote, onNominate, identity, people }) {
  const noms = [...(day.movieNominations || [])].sort((a, b) => b.votes.length - a.votes.length);
  const leader = noms[0];
  return (
    <div className="card">
      <h3>🎬 Movie + Fireplace Night</h3>
      <div className="small-muted">2 votes/person per night · nominations lock at 8pm</div>
      {leader && <div className="chip-row"><span className="chip amber">Current leader: {leader.title} ({leader.votes.length} votes)</span></div>}
      {noms.map(n => (
        <div key={n.id} className="list-item">
          <span>{n.title} <span className="small-muted">— {n.votes.length} votes</span></span>
          <button className={`btn small ${n.votes.includes(identity.personId) ? '' : 'ghost'}`} onClick={() => onVote(n.id)}>
            {n.votes.includes(identity.personId) ? '★ Voted' : '☆ Vote'}
          </button>
        </div>
      ))}
      <button className="btn small secondary" onClick={() => onNominate(day.id)}>+ Nominate a movie</button>
    </div>
  );
}

function HotTub({ day, onToggle, identity, people }) {
  return (
    <div className="card">
      <h3>♨️ Hot Tub Sign-ups</h3>
      <div className="small-muted">45-min slots, max 4 people</div>
      {day.hottubSlots.map(slot => (
        <div key={slot.id} className="list-item">
          <span>{slot.label} <span className="small-muted">({slot.signups.length}/4)</span></span>
          <button
            className={`btn small ${slot.signups.includes(identity.personId) ? '' : 'ghost'}`}
            disabled={!slot.signups.includes(identity.personId) && slot.signups.length >= 4}
            onClick={() => onToggle(slot.id)}
          >
            {slot.signups.includes(identity.personId) ? 'Leave' : 'Join'}
          </button>
        </div>
      ))}
    </div>
  );
}
