import { useState } from 'react';
import { useIdentity } from '../identity';
import { api } from '../api';

export default function Settings({ people, settings }) {
  const identity = useIdentity();
  const peopleList = people.data || [];
  const [newName, setNewName] = useState('');
  const s = settings.data || {};
  const [wifi, setWifi] = useState(s.wifiPassword || '');
  const [checkIn, setCheckIn] = useState(s.checkIn || '');
  const [checkOut, setCheckOut] = useState(s.checkOut || '');

  async function addPerson() {
    if (!newName.trim()) return;
    await api.post('/people', { name: newName.trim() });
    setNewName('');
    people.reload();
  }
  async function rename(p) {
    const name = prompt('Rename', p.name);
    if (!name) return;
    await api.patch(`/people/${p.id}`, { name });
    people.reload();
  }
  async function toggleStatus(p) {
    await api.patch(`/people/${p.id}`, { status: p.status === 'confirmed' ? 'maybe' : 'confirmed' });
    people.reload();
  }
  async function removePerson(p) {
    if (!confirm(`Remove ${p.name} from the roster?`)) return;
    await api.del(`/people/${p.id}`);
    people.reload();
  }
  async function saveHouseInfo() {
    await api.patch('/settings', { wifiPassword: wifi, checkIn, checkOut });
    settings.reload();
  }

  return (
    <div>
      <div className="card">
        <h3>👤 Roster</h3>
        {peopleList.map(p => (
          <div key={p.id} className="list-item">
            <span>{p.name}{p.isGuestOfHonor ? ' 🎂' : ''}{p.isOrganizer ? ' (organizer)' : ''}</span>
            <span className="row">
              <button className="btn small ghost" onClick={() => toggleStatus(p)}>{p.status}</button>
              <button className="btn small ghost" onClick={() => rename(p)}>Rename</button>
              {!p.isOrganizer && <button className="btn small danger" onClick={() => removePerson(p)}>×</button>}
            </span>
          </div>
        ))}
        <div className="row" style={{ marginTop: 10 }}>
          <input type="text" placeholder="Add a guest" value={newName} onChange={e => setNewName(e.target.value)} />
          <button className="btn small" onClick={addPerson}>Add</button>
        </div>
      </div>

      <div className="card">
        <h3>🏡 House Info Editable Fields</h3>
        <div className="field"><label>WiFi Password</label><input type="text" value={wifi} onChange={e => setWifi(e.target.value)} /></div>
        <div className="field"><label>Check-in</label><input type="text" value={checkIn} onChange={e => setCheckIn(e.target.value)} /></div>
        <div className="field"><label>Check-out</label><input type="text" value={checkOut} onChange={e => setCheckOut(e.target.value)} /></div>
        <button className="btn small" onClick={saveHouseInfo}>Save</button>
      </div>

      <div className="card">
        <h3>ℹ️ About You</h3>
        <p className="small-muted">Signed in as <strong>{identity.personName}</strong>.</p>
        <button className="btn small ghost" onClick={() => identity.clearPerson()}>Switch person</button>
      </div>
    </div>
  );
}
