import { useState } from 'react';
import { api } from '../api';

const STEPS = ['Trip Basics', 'House Info', 'Roster', 'Rooms', 'Review'];

const emptyPerson = () => ({ name: '', status: 'confirmed', isOrganizer: false, isGuestOfHonor: false });
const emptyRoom = () => ({ name: '', bed: '', capacity: 2, details: '' });

export default function Setup({ settings }) {
  const [step, setStep] = useState(0);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState('');

  const [trip, setTrip] = useState({
    tripName: 'Cabin Fever 2026',
    tripStart: '',
    tripEnd: '',
    address: '',
    rentalName: '',
    rentalLink: '',
    costPerPerson: '',
    lodgingCost: '',
    foodCost: '',
    payment1Label: 'Payment 1',
    payment1Date: '',
    payment1Amount: '',
    payment2Label: 'Payment 2',
    payment2Date: '',
    payment2Amount: '',
  });
  const [house, setHouse] = useState({
    checkIn: '', checkOut: '', wifiPassword: '', houseRules: '', quietHours: '', altitudeTips: '', emergencyInfo: '', houseDescription: '',
  });
  const [roster, setRoster] = useState([emptyPerson()]);
  const [rooms, setRooms] = useState([emptyRoom()]);

  function updateTrip(k, v) { setTrip(t => ({ ...t, [k]: v })); }
  function updateHouse(k, v) { setHouse(h => ({ ...h, [k]: v })); }

  function updatePerson(i, k, v) {
    setRoster(list => list.map((p, idx) => idx === i ? { ...p, [k]: v } : p));
  }
  function addPerson() { setRoster(list => [...list, emptyPerson()]); }
  function removePerson(i) { setRoster(list => list.filter((_, idx) => idx !== i)); }

  function updateRoom(i, k, v) {
    setRooms(list => list.map((r, idx) => idx === i ? { ...r, [k]: v } : r));
  }
  function addRoom() { setRooms(list => [...list, emptyRoom()]); }
  function removeRoom(i) { setRooms(list => list.filter((_, idx) => idx !== i)); }

  function canAdvance() {
    if (step === 0) return trip.tripName.trim() && trip.tripStart && trip.tripEnd;
    return true;
  }

  async function finish() {
    setSaving(true);
    setError('');
    try {
      await api.post('/setup', {
        ...trip,
        ...house,
        roster: roster.filter(p => p.name.trim()),
        rooms: rooms.filter(r => r.name.trim()),
      });
      await settings.reload();
    } catch (e) {
      setError(e.message);
      setSaving(false);
    }
  }

  return (
    <div className="setup-wizard">
      <div className="hero-emoji">🍂🏔️🔥</div>
      <h1>Let's set up your trip</h1>
      <p className="small-muted">A few quick steps — everything here is editable later in Settings.</p>

      <div className="setup-steps">
        {STEPS.map((s, i) => <span key={s} className={`setup-step ${i === step ? 'active' : i < step ? 'done' : ''}`}>{s}</span>)}
      </div>

      {step === 0 && (
        <div className="card">
          <h3>🏔️ Trip Basics</h3>
          <div className="field"><label>Trip name</label><input type="text" value={trip.tripName} onChange={e => updateTrip('tripName', e.target.value)} /></div>
          <div className="row">
            <div className="field"><label>Start date</label><input type="date" value={trip.tripStart} onChange={e => updateTrip('tripStart', e.target.value)} /></div>
            <div className="field"><label>End date</label><input type="date" value={trip.tripEnd} onChange={e => updateTrip('tripEnd', e.target.value)} /></div>
          </div>
          <div className="field"><label>Address</label><input type="text" placeholder="123 Cabin Rd, Somewhere, CO" value={trip.address} onChange={e => updateTrip('address', e.target.value)} /></div>
          <div className="row">
            <div className="field"><label>Rental name</label><input type="text" value={trip.rentalName} onChange={e => updateTrip('rentalName', e.target.value)} /></div>
            <div className="field"><label>Listing link</label><input type="text" value={trip.rentalLink} onChange={e => updateTrip('rentalLink', e.target.value)} /></div>
          </div>
          <h3 style={{ marginTop: 16 }}>💵 Cost & Payments</h3>
          <div className="row">
            <div className="field"><label>Lodging cost/person</label><input type="number" value={trip.lodgingCost} onChange={e => updateTrip('lodgingCost', e.target.value)} /></div>
            <div className="field"><label>Food cost/person</label><input type="number" value={trip.foodCost} onChange={e => updateTrip('foodCost', e.target.value)} /></div>
            <div className="field"><label>Total/person</label><input type="number" value={trip.costPerPerson} onChange={e => updateTrip('costPerPerson', e.target.value)} /></div>
          </div>
          <div className="row">
            <div className="field"><label>Payment 1 label</label><input type="text" value={trip.payment1Label} onChange={e => updateTrip('payment1Label', e.target.value)} /></div>
            <div className="field"><label>Due date</label><input type="date" value={trip.payment1Date} onChange={e => updateTrip('payment1Date', e.target.value)} /></div>
            <div className="field"><label>Amount</label><input type="number" value={trip.payment1Amount} onChange={e => updateTrip('payment1Amount', e.target.value)} /></div>
          </div>
          <div className="row">
            <div className="field"><label>Payment 2 label</label><input type="text" value={trip.payment2Label} onChange={e => updateTrip('payment2Label', e.target.value)} /></div>
            <div className="field"><label>Due date</label><input type="date" value={trip.payment2Date} onChange={e => updateTrip('payment2Date', e.target.value)} /></div>
            <div className="field"><label>Amount</label><input type="number" value={trip.payment2Amount} onChange={e => updateTrip('payment2Amount', e.target.value)} /></div>
          </div>
        </div>
      )}

      {step === 1 && (
        <div className="card">
          <h3>🏡 House Info</h3>
          <div className="row">
            <div className="field"><label>Check-in</label><input type="text" placeholder="4:00 PM" value={house.checkIn} onChange={e => updateHouse('checkIn', e.target.value)} /></div>
            <div className="field"><label>Check-out</label><input type="text" placeholder="10:00 AM" value={house.checkOut} onChange={e => updateHouse('checkOut', e.target.value)} /></div>
          </div>
          <div className="field"><label>WiFi password</label><input type="text" value={house.wifiPassword} onChange={e => updateHouse('wifiPassword', e.target.value)} /></div>
          <div className="field"><label>House rules</label><textarea value={house.houseRules} onChange={e => updateHouse('houseRules', e.target.value)} /></div>
          <div className="field"><label>Quiet hours</label><input type="text" value={house.quietHours} onChange={e => updateHouse('quietHours', e.target.value)} /></div>
          <div className="field"><label>Altitude / local tips</label><textarea value={house.altitudeTips} onChange={e => updateHouse('altitudeTips', e.target.value)} /></div>
          <div className="field"><label>Emergency info</label><textarea value={house.emergencyInfo} onChange={e => updateHouse('emergencyInfo', e.target.value)} /></div>
          <div className="field"><label>About the house</label><textarea placeholder="Bedrooms, bathrooms, amenities…" value={house.houseDescription} onChange={e => updateHouse('houseDescription', e.target.value)} /></div>
        </div>
      )}

      {step === 2 && (
        <div className="card">
          <h3>👤 Roster</h3>
          <p className="small-muted">Add everyone coming. Mark guests of honor 🎂 to unlock the birthday checklist and surprise-ideas board for them.</p>
          {roster.map((p, i) => (
            <div key={i} className="setup-row">
              <input type="text" placeholder="Name" value={p.name} onChange={e => updatePerson(i, 'name', e.target.value)} />
              <select value={p.status} onChange={e => updatePerson(i, 'status', e.target.value)}>
                <option value="confirmed">Confirmed</option>
                <option value="maybe">Maybe</option>
              </select>
              <label className="check-inline"><input type="checkbox" checked={p.isOrganizer} onChange={e => updatePerson(i, 'isOrganizer', e.target.checked)} /> Organizer</label>
              <label className="check-inline"><input type="checkbox" checked={p.isGuestOfHonor} onChange={e => updatePerson(i, 'isGuestOfHonor', e.target.checked)} /> 🎂</label>
              <button className="btn small danger" onClick={() => removePerson(i)}>×</button>
            </div>
          ))}
          <button className="btn small ghost" onClick={addPerson}>+ Add person</button>
        </div>
      )}

      {step === 3 && (
        <div className="card">
          <h3>🛏️ Rooms</h3>
          <p className="small-muted">Optional — you can also add these later in Logistics.</p>
          {rooms.map((r, i) => (
            <div key={i} className="setup-row">
              <input type="text" placeholder="Room name" value={r.name} onChange={e => updateRoom(i, 'name', e.target.value)} />
              <input type="text" placeholder="Bed type" value={r.bed} onChange={e => updateRoom(i, 'bed', e.target.value)} />
              <input type="number" placeholder="Capacity" style={{ width: 70 }} value={r.capacity} onChange={e => updateRoom(i, 'capacity', e.target.value)} />
              <input type="text" placeholder="Details" value={r.details} onChange={e => updateRoom(i, 'details', e.target.value)} />
              <button className="btn small danger" onClick={() => removeRoom(i)}>×</button>
            </div>
          ))}
          <button className="btn small ghost" onClick={addRoom}>+ Add room</button>
        </div>
      )}

      {step === 4 && (
        <div className="card">
          <h3>✅ Review</h3>
          <p><strong>{trip.tripName}</strong> — {trip.tripStart || '?'} to {trip.tripEnd || '?'}</p>
          <p className="small-muted">{trip.address || 'No address set'}</p>
          <p className="small-muted">{roster.filter(p => p.name.trim()).length} people on the roster, {rooms.filter(r => r.name.trim()).length} rooms configured.</p>
          {error && <p style={{ color: 'var(--danger, #c0392b)' }}>{error}</p>}
        </div>
      )}

      <div className="row" style={{ marginTop: 12 }}>
        {step > 0 && <button className="btn secondary" onClick={() => setStep(s => s - 1)}>Back</button>}
        {step < STEPS.length - 1 && <button className="btn" disabled={!canAdvance()} onClick={() => setStep(s => s + 1)}>Next</button>}
        {step === STEPS.length - 1 && <button className="btn" disabled={saving} onClick={finish}>{saving ? 'Setting up…' : 'Finish setup 🎉'}</button>}
      </div>
    </div>
  );
}
