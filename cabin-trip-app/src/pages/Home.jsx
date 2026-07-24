import { useIdentity } from '../identity';
import { usePoll } from '../usePoll';
import { countdownLabel, nameOf, csvToNames } from '../utils';

export default function Home({ people, settings, setTab }) {
  const identity = useIdentity();
  const days = usePoll('/days', 10000);
  const meals = usePoll('/meals', 10000);
  const moneySummary = usePoll('/money-summary', 15000);
  const payments = usePoll('/payments', 15000);
  const groceries = usePoll('/groceries', 15000);

  const s = settings.data || {};
  const dayList = days.data || [];
  const mealList = meals.data || [];
  const peopleList = people.data || [];
  const groceryList = groceries.data || [];

  const todayStr = new Date().toISOString().slice(0, 10);
  let activeDay = dayList.find(d => d.date === todayStr);
  let previewLabel = null;
  if (!activeDay) {
    const upcoming = dayList.filter(d => d.date > todayStr).sort((a, b) => a.date.localeCompare(b.date))[0];
    activeDay = upcoming || dayList[0];
    previewLabel = upcoming ? 'Next up' : 'Trip wrapped — see you next year!';
  }

  const myBlocks = dayList.flatMap(d => d.blocks || []).filter(b => b.ownerId === identity.personId);
  const myMeals = mealList.filter(m => (m.cooks || '').split(',').includes(identity.personId) || (m.cleanup || '').split(',').includes(identity.personId));
  const myGroceries = groceryList.filter(g => g.claimedBy === identity.personId);
  const myPayments = (payments.data || []).filter(p => p.personId === identity.personId && p.status !== 'confirmed');

  const unownedBlocks = dayList.flatMap(d => (d.blocks || []).map(b => ({ ...b, dayLabel: d.label }))).filter(b => !b.ownerId && !b.ownerText);
  const unownedMeals = mealList.filter(m => m.plan && !m.cooks);

  const money = moneySummary.data;
  const pct = money && money.totalNeeded ? Math.min(100, Math.round((money.collected / money.totalNeeded) * 100)) : 0;

  const upcomingInstallments = [];
  for (const r of (payments.data || [])) {
    if (!upcomingInstallments.find(i => i.dueLabel === r.dueLabel && i.dueDate === r.dueDate)) {
      upcomingInstallments.push({ dueLabel: r.dueLabel, dueDate: r.dueDate });
    }
  }

  return (
    <div>
      <div className="card" style={{ background: 'linear-gradient(160deg, #d9822b, #f0a94e)', color: 'white' }}>
        <h2 style={{ color: 'white' }}>🏔️ {countdownLabel(s.tripStart)} until we head up!</h2>
        <div className="small-muted" style={{ color: 'rgba(255,255,255,0.85)' }}>{s.tripStart} → {s.tripEnd} · {s.address}</div>
        <div className="chip-row">
          {upcomingInstallments.map(i => (
            <span key={i.dueLabel + i.dueDate} className="chip">💵 {i.dueLabel}: {countdownLabel(i.dueDate)}</span>
          ))}
        </div>
      </div>

      <div className="card">
        <div className="spread">
          <h3>{previewLabel || "Today's Vibe"}</h3>
        </div>
        {activeDay ? (
          <>
            <div className="row"><strong>{activeDay.label}</strong> — {activeDay.theme}</div>
            {(activeDay.blocks || []).slice(0, 4).map(b => (
              <div key={b.id} className="list-item">
                <span>{b.time ? `${b.time} — ` : ''}{b.title}</span>
              </div>
            ))}
            <button className="btn small ghost" onClick={() => setTab('schedule')}>See full schedule →</button>
          </>
        ) : <p className="small-muted">Loading…</p>}
      </div>

      <div className="card">
        <h3>💵 Payment Progress</h3>
        {money && (
          <>
            <div className="progress-outer"><div className="progress-inner" style={{ width: `${pct}%` }} /></div>
            <div className="small-muted">${money.collected.toFixed(2)} of ${money.totalNeeded.toFixed(2)} collected ({money.confirmedHeadcount} confirmed people)</div>
          </>
        )}
        <button className="btn small ghost" onClick={() => setTab('money')}>Manage payments →</button>
      </div>

      <div className="card">
        <h3>✅ My Responsibilities</h3>
        {myBlocks.length === 0 && myMeals.length === 0 && myGroceries.length === 0 && myPayments.length === 0 && (
          <p className="small-muted">Nothing claimed yet — go grab something on Schedule or Food!</p>
        )}
        {myBlocks.map(b => <div key={b.id} className="list-item">🔥 Leading: {b.title}</div>)}
        {myMeals.map(m => <div key={m.id} className="list-item">🍽️ Meal: {m.mealType} — {m.plan}</div>)}
        {myGroceries.map(g => <div key={g.id} className="list-item">🛒 Grocery: {g.item} ({g.qty})</div>)}
        {myPayments.map(p => <div key={p.id} className="list-item">💵 {p.dueLabel} — ${p.amount.toFixed(2)} due {p.dueDate} ({p.status})</div>)}
      </div>

      {(unownedBlocks.length > 0 || unownedMeals.length > 0) && (
        <div className="card amber-glow">
          <h3>🙋 Needs an Owner</h3>
          {unownedBlocks.map(b => <div key={b.id} className="list-item"><span>{b.dayLabel}: {b.title}</span><span className="needs-owner">unclaimed</span></div>)}
          {unownedMeals.map(m => <div key={m.id} className="list-item"><span>{m.mealType}: {m.plan}</span><span className="needs-owner">needs a cook</span></div>)}
        </div>
      )}
    </div>
  );
}
