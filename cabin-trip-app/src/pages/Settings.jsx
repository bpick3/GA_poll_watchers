import { useState } from 'react';
import { useIdentity } from '../identity';
import { api } from '../api';

export default function Settings({ people, settings }) {
  const identity = useIdentity();
  const peopleList = people.data || [];
  const [newName, setNewName] = useState('');
  const s = settings.data || {};

  const [trip, setTrip] = useState({
    tripName: s.tripName || '', tripStart: s.tripStart || '', tripEnd: s.tripEnd || '',
    address: s.address || '', rentalName: s.rentalName || '', rentalLink: s.rentalLink || '',
    costPerPerson: s.costPerPerson || '', lodgingCost: s.lodgingCost || '', foodCost: s.foodCost || '',
  });
  const [house, setHouse] = useState({
    wifiPassword: s.wifiPassword || '', checkIn: s.checkIn || '', checkOut: s.checkOut || '',
    houseRules: s.houseRules || '', quietHours: s.quietHours || '', altitudeTips: s.altitudeTips || '',
    emergencyInfo: s.emergencyInfo || '', houseDescription: s.houseDescription || '',
  });

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
  async function toggleOrganizer(p) {
    await api.patch(`/people/${p.id}`, { isOrganizer: p.isOrganizer ? 0 : 1 });
    people.reload();
  }
  async function toggleGoh(p) {
    await api.patch(`/people/${p.id}`, { isGuestOfHonor: p.isGuestOfHonor ? 0 : 1 });
    people.reload();
  }
  async function removePerson(p) {
    if (!confirm(`Remove ${p.name} from the roster?`)) return;
    await api.del(`/people/${p.id}`);
    people.reload();
  }
  async function saveTrip() {
    await api.patch('/settings', trip);
    settings.reload();
  }
  async function saveHouseInfo() {
    await api.patch('/settings', house);
    settings.reload();
  }

  return (
    <div>
      <div className="card">
        <h3>🏔️ Trip Basics</h3>
        <div className="field"><label>Trip name</label><input type="text" value={trip.tripName} onChange={e => setTrip(t => ({ ...t, tripName: e.target.value }))} /></div>
        <div className="row">
          <div className="field"><label>Start date</label><input type="date" value={trip.tripStart} onChange={e => setTrip(t => ({ ...t, tripStart: e.target.value }))} /></div>
          <div className="field"><label>End date</label><input type="date" value={trip.tripEnd} onChange={e => setTrip(t => ({ ...t, tripEnd: e.target.value }))} /></div>
        </div>
        <div className="field"><label>Address</label><input type="text" value={trip.address} onChange={e => setTrip(t => ({ ...t, address: e.target.value }))} /></div>
        <div className="row">
          <div className="field"><label>Rental name</label><input type="text" value={trip.rentalName} onChange={e => setTrip(t => ({ ...t, rentalName: e.target.value }))} /></div>
          <div className="field"><label>Listing link</label><input type="text" value={trip.rentalLink} onChange={e => setTrip(t => ({ ...t, rentalLink: e.target.value }))} /></div>
        </div>
        <div className="row">
          <div className="field"><label>Lodging/person</label><input type="number" value={trip.lodgingCost} onChange={e => setTrip(t => ({ ...t, lodgingCost: e.target.value }))} /></div>
          <div className="field"><label>Food/person</label><input type="number" value={trip.foodCost} onChange={e => setTrip(t => ({ ...t, foodCost: e.target.value }))} /></div>
          <div className="field"><label>Total/person</label><input type="number" value={trip.costPerPerson} onChange={e => setTrip(t => ({ ...t, costPerPerson: e.target.value }))} /></div>
        </div>
        <button className="btn small" onClick={saveTrip}>Save trip basics</button>
      </div>

      <div className="card">
        <h3>👤 Roster</h3>
        {peopleList.map(p => (
          <div key={p.id} className="list-item">
            <span>{p.name}{p.isGuestOfHonor ? ' 🎂' : ''}{p.isOrganizer ? ' (organizer)' : ''}</span>
            <span className="row wrap">
              <button className="btn small ghost" onClick={() => toggleStatus(p)}>{p.status}</button>
              <button className="btn small ghost" onClick={() => toggleOrganizer(p)}>{p.isOrganizer ? 'Unmake organizer' : 'Make organizer'}</button>
              <button className="btn small ghost" onClick={() => toggleGoh(p)}>{p.isGuestOfHonor ? 'Remove 🎂' : 'Mark 🎂'}</button>
              <button className="btn small ghost" onClick={() => rename(p)}>Rename</button>
              <button className="btn small danger" onClick={() => removePerson(p)}>×</button>
            </span>
          </div>
        ))}
        <div className="row" style={{ marginTop: 10 }}>
          <input type="text" placeholder="Add a guest" value={newName} onChange={e => setNewName(e.target.value)} />
          <button className="btn small" onClick={addPerson}>Add</button>
        </div>
      </div>

      <div className="card">
        <h3>🏡 House Info</h3>
        <div className="row">
          <div className="field"><label>Check-in</label><input type="text" value={house.checkIn} onChange={e => setHouse(h => ({ ...h, checkIn: e.target.value }))} /></div>
          <div className="field"><label>Check-out</label><input type="text" value={house.checkOut} onChange={e => setHouse(h => ({ ...h, checkOut: e.target.value }))} /></div>
        </div>
        <div className="field"><label>WiFi Password</label><input type="text" value={house.wifiPassword} onChange={e => setHouse(h => ({ ...h, wifiPassword: e.target.value }))} /></div>
        <div className="field"><label>House Rules</label><textarea value={house.houseRules} onChange={e => setHouse(h => ({ ...h, houseRules: e.target.value }))} /></div>
        <div className="field"><label>Quiet Hours</label><input type="text" value={house.quietHours} onChange={e => setHouse(h => ({ ...h, quietHours: e.target.value }))} /></div>
        <div className="field"><label>Altitude / Local Tips</label><textarea value={house.altitudeTips} onChange={e => setHouse(h => ({ ...h, altitudeTips: e.target.value }))} /></div>
        <div className="field"><label>Emergency Info</label><textarea value={house.emergencyInfo} onChange={e => setHouse(h => ({ ...h, emergencyInfo: e.target.value }))} /></div>
        <div className="field"><label>About the House</label><textarea value={house.houseDescription} onChange={e => setHouse(h => ({ ...h, houseDescription: e.target.value }))} /></div>
        <button className="btn small" onClick={saveHouseInfo}>Save house info</button>
      </div>

      <ThemeLibrary settings={settings} />

      <div className="card">
        <h3>ℹ️ About You</h3>
        <p className="small-muted">Signed in as <strong>{identity.personName}</strong>.</p>
        <button className="btn small ghost" onClick={() => identity.clearPerson()}>Switch person</button>
      </div>
    </div>
  );
}

function ThemeLibrary({ settings }) {
  const s = settings.data || {};
  const themes = (() => {
    try { return JSON.parse(s.altThemes || '[]'); } catch { return []; }
  })();
  const [newTheme, setNewTheme] = useState('');

  async function save(next) {
    await api.patch('/settings', { altThemes: JSON.stringify(next) });
    settings.reload();
  }
  async function add() {
    if (!newTheme.trim()) return;
    await save([...themes, newTheme.trim()]);
    setNewTheme('');
  }
  async function remove(t) {
    await save(themes.filter(x => x !== t));
  }

  return (
    <div className="card">
      <h3>🎭 Day Theme Library</h3>
      <p className="small-muted">These show up as quick-pick options when changing a day's theme on the Schedule tab.</p>
      {themes.map(t => (
        <div key={t} className="list-item">
          <span>{t}</span>
          <button className="btn small danger" onClick={() => remove(t)}>×</button>
        </div>
      ))}
      <div className="row" style={{ marginTop: 10 }}>
        <input type="text" placeholder="e.g. Board Game Bonanza 🎲" value={newTheme} onChange={e => setNewTheme(e.target.value)} />
        <button className="btn small" onClick={add}>Add</button>
      </div>
    </div>
  );
}
