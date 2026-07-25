import { useState } from 'react';
import { useIdentity } from '../identity';
import { usePoll } from '../usePoll';
import { api } from '../api';
import { nameOf } from '../utils';

export default function Games({ people, settings }) {
  const identity = useIdentity();
  const [sub, setSub] = useState('library');
  const bring = usePoll('/games-bring', 8000);
  const tournaments = usePoll('/tournaments', 8000);
  const birthdayChecklist = usePoll('/birthday-checklist', 8000);
  const surprises = usePoll('/surprise-ideas', 8000);
  const peopleList = people.data || [];
  const guestsOfHonor = peopleList.filter(p => p.isGuestOfHonor);
  const [newGame, setNewGame] = useState('');
  const [newHouseGame, setNewHouseGame] = useState('');

  const isGuestOfHonor = guestsOfHonor.some(p => p.id === identity.personId);

  const houseGames = (() => {
    try { return JSON.parse((settings?.data || {}).houseGames || '[]'); } catch { return []; }
  })();

  async function claimGame(g) {
    const mine = identity.personName;
    await api.patch(`/games-bring/${g.id}`, { claimedBy: g.claimedBy ? null : mine });
    bring.reload();
  }
  async function addBringGame() {
    if (!newGame.trim()) return;
    await api.post('/games-bring', { name: newGame.trim() });
    setNewGame('');
    bring.reload();
  }
  async function removeBringGame(id) {
    await api.del(`/games-bring/${id}`);
    bring.reload();
  }
  async function addHouseGame() {
    if (!newHouseGame.trim()) return;
    await api.patch('/settings', { houseGames: JSON.stringify([...houseGames, newHouseGame.trim()]) });
    setNewHouseGame('');
    settings.reload();
  }
  async function removeHouseGame(g) {
    await api.patch('/settings', { houseGames: JSON.stringify(houseGames.filter(x => x !== g)) });
    settings.reload();
  }

  return (
    <div>
      <div className="subnav">
        <button className={sub === 'library' ? 'active' : ''} onClick={() => setSub('library')}>Game Library</button>
        <button className={sub === 'tournament' ? 'active' : ''} onClick={() => setSub('tournament')}>Tournament</button>
        <button className={sub === 'birthday' ? 'active' : ''} onClick={() => setSub('birthday')}>🎂 Birthday</button>
      </div>

      {sub === 'library' && (
        <div className="card">
          <h3>🏠 House Games</h3>
          <div className="chip-row">
            {houseGames.map(g => <span key={g} className="chip" onClick={() => removeHouseGame(g)} style={{ cursor: 'pointer' }} title="Tap to remove">{g} ×</span>)}
            {houseGames.length === 0 && <span className="small-muted">None added yet — what's already at the house? (ping pong, foosball, etc.)</span>}
          </div>
          <div className="row" style={{ marginTop: 8 }}>
            <input type="text" placeholder="e.g. 🏓 Ping Pong" value={newHouseGame} onChange={e => setNewHouseGame(e.target.value)} />
            <button className="btn small" onClick={addHouseGame}>Add</button>
          </div>

          <h3 style={{ marginTop: 16 }}>🎲 Bring List</h3>
          {(bring.data || []).map(g => (
            <div key={g.id} className="list-item">
              <span>{g.name}</span>
              <span className="row">
                <button className={`btn small ${g.claimedBy ? '' : 'ghost'}`} onClick={() => claimGame(g)}>{g.claimedBy || 'Claim'}</button>
                <button className="btn small danger" onClick={() => removeBringGame(g.id)}>×</button>
              </span>
            </div>
          ))}
          <div className="row" style={{ marginTop: 8 }}>
            <input type="text" placeholder="e.g. Codenames" value={newGame} onChange={e => setNewGame(e.target.value)} />
            <button className="btn small" onClick={addBringGame}>Add</button>
          </div>
        </div>
      )}

      {sub === 'tournament' && <Tournament tournaments={tournaments} people={peopleList} />}

      {sub === 'birthday' && (
        <div>
          {guestsOfHonor.length === 0 && <p className="small-muted">No guests of honor set — mark someone 🎂 in Settings to unlock birthday planning.</p>}
          {isGuestOfHonor && (
            <div className="card amber-glow">
              <p className="small-muted">🤫 Your own birthday card is hidden from you — everyone else can still see and plan it. You can see everyone else's below.</p>
            </div>
          )}
          {guestsOfHonor.filter(person => person.id !== identity.personId).map(person => (
            <BirthdayCard key={person.id} person={person} checklist={birthdayChecklist} />
          ))}
          {!isGuestOfHonor && <SurpriseIdeas surprises={surprises} people={peopleList} identity={identity} />}
        </div>
      )}
    </div>
  );
}

function BirthdayCard({ person, checklist }) {
  const [item, setItem] = useState('');
  const items = (checklist.data || []).filter(i => i.forPerson === person.name);

  async function toggle(i) {
    await api.patch(`/birthday-checklist/${i.id}`, { done: !i.done });
    checklist.reload();
  }
  async function remove(id) {
    await api.del(`/birthday-checklist/${id}`);
    checklist.reload();
  }
  async function add() {
    if (!item.trim()) return;
    await api.post('/birthday-checklist', { forPerson: person.name, item: item.trim() });
    setItem('');
    checklist.reload();
  }

  return (
    <div className="card">
      <h3>🎂 {person.name}'s Moments</h3>
      {items.map(i => (
        <div key={i.id} className={`checklist-item ${i.done ? 'done' : ''}`}>
          <input type="checkbox" checked={!!i.done} onChange={() => toggle(i)} />
          <span style={{ flex: 1 }}>{i.item}</span>
          <button className="btn small ghost" onClick={() => remove(i.id)}>×</button>
        </div>
      ))}
      <div className="row" style={{ marginTop: 8 }}>
        <input type="text" placeholder="Add a to-do…" value={item} onChange={e => setItem(e.target.value)} />
        <button className="btn small" onClick={add}>Add</button>
      </div>
    </div>
  );
}

function SurpriseIdeas({ surprises, people, identity }) {
  const [text, setText] = useState('');
  async function add() {
    if (!text.trim()) return;
    await api.post('/surprise-ideas', { text });
    setText('');
    surprises.reload();
  }
  return (
    <div className="card amber-glow">
      <h3>🤫 Surprise Ideas</h3>
      <div className="small-muted">Hidden from the guests of honor.</div>
      {(surprises.data || []).map(s => (
        <div key={s.id} className="note-bubble">
          <div className="meta">{nameOf(people, s.personId) || 'Someone'}</div>
          {s.text}
        </div>
      ))}
      <div className="field" style={{ marginTop: 8 }}>
        <textarea value={text} onChange={e => setText(e.target.value)} placeholder="Got an idea?" />
      </div>
      <button className="btn small" onClick={add}>Post idea</button>
    </div>
  );
}

function Tournament({ tournaments, people }) {
  const list = tournaments.data || [];
  const t = list[0];
  const [name, setName] = useState('Tournament');
  const [selected, setSelected] = useState([]);

  async function advance(matchId, winner) {
    await api.post(`/tournaments/${t.id}/advance`, { matchId, winner });
    tournaments.reload();
  }

  function toggleSelected(id) {
    setSelected(sel => sel.includes(id) ? sel.filter(x => x !== id) : [...sel, id]);
  }

  async function create() {
    if (selected.length < 2) return alert('Pick at least 2 players.');
    const players = people.filter(p => selected.includes(p.id)).map(p => p.name);
    await api.post('/tournaments', { name, players });
    tournaments.reload();
  }

  if (!t) {
    return (
      <div className="card">
        <h3>🏓 New Tournament</h3>
        <div className="field"><label>Name</label><input type="text" value={name} onChange={e => setName(e.target.value)} /></div>
        <div className="field">
          <label>Players</label>
          <div className="chip-row">
            {people.map(p => (
              <span key={p.id} className={`chip ${selected.includes(p.id) ? 'amber' : ''}`} style={{ cursor: 'pointer' }} onClick={() => toggleSelected(p.id)}>{p.name}</span>
            ))}
          </div>
        </div>
        <button className="btn small" onClick={create}>Create bracket</button>
      </div>
    );
  }

  return (
    <div className="card">
      <h3>🏓 {t.name}</h3>
      {t.championId && <div className="chip-row"><span className="chip amber">👑 Champion: {t.championId}</span></div>}
      {t.bracket.rounds.map((round, ri) => (
        <div key={ri} className="bracket-round">
          <div className="small-muted"><strong>Round {ri + 1}</strong></div>
          {round.map(m => (
            <div key={m.id} className="bracket-match">
              <div className="small-muted">{m.p1} vs {m.p2 || 'BYE'}</div>
              {m.p1 && <button className={m.winner === m.p1 ? 'winner' : ''} onClick={() => advance(m.id, m.p1)}>{m.p1}{m.winner === m.p1 ? ' 👑' : ''}</button>}
              {m.p2 && <button className={m.winner === m.p2 ? 'winner' : ''} onClick={() => advance(m.id, m.p2)}>{m.p2}{m.winner === m.p2 ? ' 👑' : ''}</button>}
            </div>
          ))}
        </div>
      ))}
    </div>
  );
}
