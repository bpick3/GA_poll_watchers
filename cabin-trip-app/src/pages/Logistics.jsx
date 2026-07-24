import { useState } from 'react';
import { useIdentity } from '../identity';
import { usePoll } from '../usePoll';
import { api } from '../api';
import { nameOf } from '../utils';

export default function Logistics({ people, settings }) {
  const [sub, setSub] = useState('house');
  const rooms = usePoll('/rooms', 8000);
  const carpools = usePoll('/carpools', 8000);
  const packing = usePoll('/packing', 8000);
  const cleanup = usePoll('/cleanup-tasks', 8000);
  const peopleList = people.data || [];

  return (
    <div>
      <div className="subnav">
        <button className={sub === 'house' ? 'active' : ''} onClick={() => setSub('house')}>House Info</button>
        <button className={sub === 'rooms' ? 'active' : ''} onClick={() => setSub('rooms')}>Rooms</button>
        <button className={sub === 'carpool' ? 'active' : ''} onClick={() => setSub('carpool')}>Carpool</button>
        <button className={sub === 'packing' ? 'active' : ''} onClick={() => setSub('packing')}>Packing</button>
        <button className={sub === 'cleanup' ? 'active' : ''} onClick={() => setSub('cleanup')}>Cleanup</button>
      </div>
      {sub === 'house' && <HouseInfo settings={settings} />}
      {sub === 'rooms' && <Rooms rooms={rooms} people={peopleList} />}
      {sub === 'carpool' && <Carpool carpools={carpools} people={peopleList} />}
      {sub === 'packing' && <Packing packing={packing} />}
      {sub === 'cleanup' && <Cleanup cleanup={cleanup} />}
    </div>
  );
}

function HouseInfo({ settings }) {
  const s = settings.data || {};
  const mapsUrl = `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(s.address || '')}`;
  return (
    <div>
      <div className="card">
        <h3>🏡 {s.address || 'House Info'}</h3>
        <a href={mapsUrl} target="_blank" rel="noreferrer" className="btn small secondary" style={{ display: 'inline-block', marginBottom: 8 }}>📍 Open in Google Maps</a>
        <div className="small-muted">{s.rentalName} — <a href={s.rentalLink} target="_blank" rel="noreferrer">listing link</a></div>
        <div className="chip-row">
          <span className="chip">Check-in {s.checkIn}</span>
          <span className="chip">Check-out {s.checkOut}</span>
        </div>
        <div className="field" style={{ marginTop: 10 }}><label>WiFi Password</label><div className="pill">{s.wifiPassword}</div></div>
        <div className="field"><label>House Rules</label><p>{s.houseRules}</p></div>
        <div className="field"><label>Quiet Hours</label><p>{s.quietHours}</p></div>
        <div className="field"><label>Altitude Tips 🏔️</label><p>{s.altitudeTips}</p></div>
        <div className="field"><label>Emergency Info</label><p>{s.emergencyInfo}</p></div>
      </div>
      {s.houseDescription && (
        <div className="card">
          <h3>The House</h3>
          <p className="small-muted">{s.houseDescription}</p>
        </div>
      )}
    </div>
  );
}

function Rooms({ rooms, people }) {
  const list = rooms.data || [];
  const [name, setName] = useState('');
  const [bed, setBed] = useState('');
  const [capacity, setCapacity] = useState(2);
  const [details, setDetails] = useState('');

  async function addRoom() {
    if (!name.trim()) return;
    await api.post('/rooms', { name, bed, capacity: Number(capacity), details });
    setName(''); setBed(''); setDetails('');
    rooms.reload();
  }
  async function removeRoom(id) {
    if (!confirm('Remove this room?')) return;
    await api.del(`/rooms/${id}`);
    rooms.reload();
  }
  async function assign(roomId, personId) {
    try {
      await api.post(`/rooms/${roomId}/assign`, { personId });
      rooms.reload();
    } catch (e) { alert(e.message); }
  }
  async function unassign(roomId, personId) {
    await api.post(`/rooms/${roomId}/unassign`, { personId });
    rooms.reload();
  }
  async function setNote(personId, note) {
    await api.patch(`/room-assignments/${personId}/note`, { note });
    rooms.reload();
  }
  const assignedIds = list.flatMap(r => r.occupants.map(o => o.personId));
  const unassigned = people.filter(p => !assignedIds.includes(p.id));

  return (
    <div>
      {list.map(r => (
        <div key={r.id} className="card room-card">
          <span className="cap">{r.occupants.length}/{r.capacity}</span>
          <h3>{r.name}</h3>
          <div className="small-muted">{r.bed} · {r.details}</div>
          <button className="btn small ghost" onClick={() => removeRoom(r.id)}>Remove room</button>
          {r.occupants.map(o => (
            <div key={o.personId} className="list-item">
              <span>{nameOf(people, o.personId)}</span>
              <button className="btn small ghost" onClick={() => unassign(r.id, o.personId)}>Remove</button>
            </div>
          ))}
          {r.occupants.length < r.capacity && (
            <select defaultValue="" onChange={e => e.target.value && assign(r.id, e.target.value)}>
              <option value="">+ Assign someone…</option>
              {unassigned.map(p => <option key={p.id} value={p.id}>{p.name}</option>)}
            </select>
          )}
        </div>
      ))}
      <div className="card">
        <h3>+ Add a room</h3>
        <div className="field"><label>Name</label><input type="text" value={name} onChange={e => setName(e.target.value)} /></div>
        <div className="row">
          <div className="field"><label>Bed type</label><input type="text" value={bed} onChange={e => setBed(e.target.value)} /></div>
          <div className="field"><label>Capacity</label><input type="number" style={{ width: 80 }} value={capacity} onChange={e => setCapacity(e.target.value)} /></div>
        </div>
        <div className="field"><label>Details</label><input type="text" value={details} onChange={e => setDetails(e.target.value)} /></div>
        <button className="btn small" onClick={addRoom}>Add room</button>
      </div>
    </div>
  );
}

function Carpool({ carpools, people }) {
  const list = carpools.data || [];
  const [driverId, setDriverId] = useState('');
  const [seats, setSeats] = useState(3);
  const [departure, setDeparture] = useState('');
  const [eta, setEta] = useState('');

  async function add() {
    if (!driverId) return;
    await api.post('/carpools', { driverId, seats: Number(seats), departure, eta });
    setDriverId(''); setDeparture(''); setEta('');
    carpools.reload();
  }
  async function claim(id) {
    try {
      await api.post(`/carpools/${id}/claim`, {});
      carpools.reload();
    } catch (e) { alert(e.message); }
  }

  return (
    <div>
      {list.map(c => (
        <div key={c.id} className="card">
          <h3>🚗 {nameOf(people, c.driverId)} driving</h3>
          <div className="small-muted">Leaves {c.departure} · ETA {c.eta} · {c.riders.length}/{c.seats} seats filled</div>
          <div className="row wrap">{c.riders.map(rid => <span key={rid} className="pill">{nameOf(people, rid)}</span>)}</div>
          <button className="btn small ghost" style={{ marginTop: 8 }} onClick={() => claim(c.id)}>Claim / Release a seat</button>
        </div>
      ))}
      <div className="card">
        <h3>+ Offer a ride</h3>
        <div className="field"><label>Driver</label><select value={driverId} onChange={e => setDriverId(e.target.value)}><option value="">Select…</option>{people.map(p => <option key={p.id} value={p.id}>{p.name}</option>)}</select></div>
        <div className="field"><label>Seats</label><input type="number" value={seats} onChange={e => setSeats(e.target.value)} /></div>
        <div className="field"><label>Departure</label><input type="text" placeholder="e.g. Denver, 12pm Fri" value={departure} onChange={e => setDeparture(e.target.value)} /></div>
        <div className="field"><label>ETA</label><input type="text" value={eta} onChange={e => setEta(e.target.value)} /></div>
        <button className="btn small" onClick={add}>Add ride</button>
      </div>
    </div>
  );
}

function Packing({ packing }) {
  const list = packing.data || [];
  const identity = useIdentity();
  const [item, setItem] = useState('');

  async function toggle(id) {
    await api.post(`/packing/${id}/toggle`, {});
    packing.reload();
  }
  async function add() {
    if (!item.trim()) return;
    await api.post('/packing', { item });
    setItem('');
    packing.reload();
  }

  return (
    <div className="card">
      <h3>🎒 Shared Packing List</h3>
      <div className="small-muted">Tap to check off your personal copy — altitude & October essentials.</div>
      {list.map(i => {
        const mine = i.checks.find(c => c.personId === identity.personId);
        return (
          <div key={i.id} className={`checklist-item ${mine?.checked ? 'done' : ''}`}>
            <input type="checkbox" checked={!!mine?.checked} onChange={() => toggle(i.id)} />
            <span>{i.item}</span>
          </div>
        );
      })}
      <div className="row" style={{ marginTop: 10 }}>
        <input type="text" placeholder="Add item" value={item} onChange={e => setItem(e.target.value)} />
        <button className="btn small" onClick={add}>Add</button>
      </div>
    </div>
  );
}

function Cleanup({ cleanup }) {
  const list = cleanup.data || [];
  const identity = useIdentity();
  async function claim(t) {
    await api.patch(`/cleanup-tasks/${t.id}`, { claimedBy: t.claimedBy === identity.personName ? null : identity.personName, done: t.done });
    cleanup.reload();
  }
  async function toggleDone(t) {
    await api.patch(`/cleanup-tasks/${t.id}`, { claimedBy: t.claimedBy, done: !t.done });
    cleanup.reload();
  }
  return (
    <div className="card">
      <h3>🧹 Checkout Cleanup</h3>
      {list.map(t => (
        <div key={t.id} className={`checklist-item ${t.done ? 'done' : ''}`}>
          <input type="checkbox" checked={!!t.done} onChange={() => toggleDone(t)} />
          <span style={{ flex: 1 }}>{t.task}</span>
          <button className="btn small ghost" onClick={() => claim(t)}>{t.claimedBy || 'Claim'}</button>
        </div>
      ))}
    </div>
  );
}
