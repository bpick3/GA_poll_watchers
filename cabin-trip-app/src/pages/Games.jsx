import { useState } from 'react';
import { useIdentity } from '../identity';
import { usePoll } from '../usePoll';
import { api } from '../api';
import { nameOf } from '../utils';

export default function Games({ people }) {
  const identity = useIdentity();
  const [sub, setSub] = useState('library');
  const bring = usePoll('/games-bring', 8000);
  const tournaments = usePoll('/tournaments', 8000);
  const birthdayChecklist = usePoll('/birthday-checklist', 8000);
  const surprises = usePoll('/surprise-ideas', 8000);
  const peopleList = people.data || [];

  const isGuestOfHonor = ['Brandon', 'Rachel', 'Lance'].includes(identity.personName);

  async function claimGame(g) {
    const mine = identity.personName;
    await api.patch(`/games-bring/${g.id}`, { claimedBy: g.claimedBy ? null : mine });
    bring.reload();
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
          <div className="chip-row"><span className="chip">🏓 Ping Pong</span><span className="chip">⚽ Foosball</span><span className="chip">🕹️ Pinball</span></div>
          <h3 style={{ marginTop: 16 }}>🎲 Bring List</h3>
          {(bring.data || []).map(g => (
            <div key={g.id} className="list-item">
              <span>{g.name}</span>
              <button className={`btn small ${g.claimedBy ? '' : 'ghost'}`} onClick={() => claimGame(g)}>{g.claimedBy || 'Claim'}</button>
            </div>
          ))}
        </div>
      )}

      {sub === 'tournament' && <Tournament tournaments={tournaments} people={peopleList} />}

      {sub === 'birthday' && (
        <div>
          {['Brandon', 'Rachel', 'Lance'].map(person => (
            <div key={person} className="card">
              <h3>🎂 {person}'s Moments</h3>
              {(birthdayChecklist.data || []).filter(i => i.forPerson === person).map(i => (
                <div key={i.id} className={`checklist-item ${i.done ? 'done' : ''}`}>
                  <input type="checkbox" checked={!!i.done} onChange={async () => { await api.patch(`/birthday-checklist/${i.id}`, { done: !i.done }); birthdayChecklist.reload(); }} />
                  <span>{i.item}</span>
                </div>
              ))}
            </div>
          ))}
          {!isGuestOfHonor && <SurpriseIdeas surprises={surprises} people={peopleList} identity={identity} />}
          {isGuestOfHonor && (
            <div className="card">
              <p className="small-muted">🤫 This part of the app is hidden from birthday people. Nice try!</p>
            </div>
          )}
        </div>
      )}
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
      <div className="small-muted">Hidden from Brandon, Rachel & Lance.</div>
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

  async function advance(matchId, winner) {
    await api.post(`/tournaments/${t.id}/advance`, { matchId, winner });
    tournaments.reload();
  }

  if (!t) return <p className="small-muted">No tournament yet.</p>;

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
