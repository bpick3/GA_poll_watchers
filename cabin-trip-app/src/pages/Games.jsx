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
        <button className={sub === 'teams' ? 'active' : ''} onClick={() => setSub('teams')}>Teams</button>
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

      {sub === 'teams' && <Teams people={peopleList} identity={identity} />}

      {sub === 'tournament' && <Tournament tournaments={tournaments} people={peopleList} identity={identity} />}

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

function Teams({ people, identity }) {
  const teams = usePoll('/teams', 8000);
  const list = teams.data || [];
  const [newName, setNewName] = useState('');

  async function create() {
    if (!newName.trim()) return;
    await api.post('/teams', { name: newName.trim() });
    setNewName('');
    teams.reload();
  }
  async function remove(id) {
    if (!confirm('Delete this team?')) return;
    await api.del(`/teams/${id}`);
    teams.reload();
  }
  async function toggleJoin(team) {
    const inIt = team.members.includes(identity.personId);
    await api.post(`/teams/${team.id}/${inIt ? 'leave' : 'join'}`, {});
    teams.reload();
  }

  return (
    <div>
      {list.map(team => (
        <div key={team.id} className="card">
          <div className="spread">
            <h3>🎽 {team.name}</h3>
            <span className="row">
              <button className={`btn small ${team.members.includes(identity.personId) ? '' : 'ghost'}`} onClick={() => toggleJoin(team)}>
                {team.members.includes(identity.personId) ? 'Leave' : 'Join'}
              </button>
              <button className="btn small danger" onClick={() => remove(team.id)}>×</button>
            </span>
          </div>
          <div className="small-muted">{team.members.map(id => nameOf(people, id)).join(', ') || 'No members yet'}</div>
        </div>
      ))}
      <div className="card">
        <h3>+ New Team</h3>
        <div className="row">
          <input type="text" placeholder="Team name" value={newName} onChange={e => setNewName(e.target.value)} />
          <button className="btn small" onClick={create}>Add</button>
        </div>
      </div>
    </div>
  );
}

function Tournament({ tournaments, people, identity }) {
  const teamsPoll = usePoll('/teams', 8000);
  const teamList = teamsPoll.data || [];
  const list = tournaments.data || [];
  const t = list[0];
  const [name, setName] = useState('Tournament');
  const [mode, setMode] = useState('individual');
  const [selected, setSelected] = useState([]);

  const me = people.find(p => p.id === identity.personId);
  const isOrganizer = !!me?.isOrganizer;

  async function vote(matchId, choice) {
    await api.post(`/tournaments/${t.id}/vote`, { matchId, choice });
    tournaments.reload();
  }

  async function houseRuling(match, roundIdx, winnerName) {
    const isLastRound = t && roundIdx === t.bracket.rounds.length - 1;
    if (!isLastRound && !confirm(`House ruling: ${winnerName} won. This clears any later round already built off this match. Continue?`)) return;
    await api.post(`/tournaments/${t.id}/advance`, { matchId: match.id, winner: winnerName });
    tournaments.reload();
  }

  function toggleSelected(id) {
    setSelected(sel => sel.includes(id) ? sel.filter(x => x !== id) : [...sel, id]);
  }

  async function create() {
    if (selected.length < 2) return alert(mode === 'team' ? 'Pick at least 2 teams.' : 'Pick at least 2 players.');
    const source = mode === 'team' ? teamList : people;
    const players = source.filter(x => selected.includes(x.id)).map(x => x.name);
    await api.post('/tournaments', { name, players, mode });
    tournaments.reload();
  }

  async function removeTournament() {
    if (!confirm(`Delete "${t.name}"? This can't be undone.`)) return;
    await api.del(`/tournaments/${t.id}`);
    tournaments.reload();
  }

  if (!t) {
    const source = mode === 'team' ? teamList : people;
    return (
      <div className="card">
        <h3>🏓 New Tournament</h3>
        <div className="field"><label>Name</label><input type="text" value={name} onChange={e => setName(e.target.value)} /></div>
        <div className="field">
          <label>Format</label>
          <div className="row">
            <button className={`btn small ${mode === 'individual' ? '' : 'ghost'}`} onClick={() => { setMode('individual'); setSelected([]); }}>Individual</button>
            <button className={`btn small ${mode === 'team' ? '' : 'ghost'}`} onClick={() => { setMode('team'); setSelected([]); }}>Team</button>
          </div>
        </div>
        <div className="field">
          <label>{mode === 'team' ? 'Teams' : 'Players'}</label>
          {mode === 'team' && teamList.length === 0 && <p className="small-muted">No teams yet — add some on the Teams tab first.</p>}
          <div className="chip-row">
            {source.map(x => (
              <span key={x.id} className={`chip ${selected.includes(x.id) ? 'amber' : ''}`} style={{ cursor: 'pointer' }} onClick={() => toggleSelected(x.id)}>{x.name}</span>
            ))}
          </div>
        </div>
        <button className="btn small" onClick={create}>Create bracket</button>
      </div>
    );
  }

  return (
    <div className="card">
      <div className="spread">
        <h3>🏓 {t.name}</h3>
        <button className="btn small danger" onClick={removeTournament}>Delete</button>
      </div>
      <div className="small-muted">Vote for who actually won each match — the winner is whichever side gets more votes. Tap your pick again to take back your vote.</div>
      {t.championId && <div className="chip-row"><span className="chip amber">👑 Champion: {t.championId}</span></div>}
      {t.bracket.rounds.map((round, ri) => (
        <div key={ri} className="bracket-round">
          <div className="small-muted"><strong>Round {ri + 1}</strong></div>
          {round.map(m => {
            const votes = m.votes || {};
            const p1Votes = Object.values(votes).filter(v => v === 'p1').length;
            const p2Votes = Object.values(votes).filter(v => v === 'p2').length;
            const myVote = votes[identity.personId];
            const isBye = !m.p1 || !m.p2;
            return (
              <div key={m.id} className="bracket-match">
                <div className="small-muted">{m.p1} vs {m.p2 || 'BYE'}</div>
                {isBye ? (
                  <div className="small-muted">{m.winner} advances automatically</div>
                ) : (
                  <>
                    <button className={`${m.winner === m.p1 ? 'winner' : ''} ${myVote === 'p1' ? 'voted' : ''}`} onClick={() => vote(m.id, 'p1')}>
                      {m.p1} — {p1Votes} vote{p1Votes === 1 ? '' : 's'}{m.winner === m.p1 ? ' 👑' : ''}{myVote === 'p1' ? ' ✓' : ''}
                    </button>
                    <button className={`${m.winner === m.p2 ? 'winner' : ''} ${myVote === 'p2' ? 'voted' : ''}`} onClick={() => vote(m.id, 'p2')}>
                      {m.p2} — {p2Votes} vote{p2Votes === 1 ? '' : 's'}{m.winner === m.p2 ? ' 👑' : ''}{myVote === 'p2' ? ' ✓' : ''}
                    </button>
                    {p1Votes > 0 && p1Votes === p2Votes && <div className="small-muted">Tied — needs one more vote to break the tie.</div>}
                    {isOrganizer && (
                      <div className="row" style={{ marginTop: 6 }}>
                        <span className="small-muted">🔨 House Ruling:</span>
                        <button className="btn small ghost" onClick={() => houseRuling(m, ri, m.p1)}>{m.p1}</button>
                        <button className="btn small ghost" onClick={() => houseRuling(m, ri, m.p2)}>{m.p2}</button>
                      </div>
                    )}
                  </>
                )}
              </div>
            );
          })}
        </div>
      ))}
    </div>
  );
}
