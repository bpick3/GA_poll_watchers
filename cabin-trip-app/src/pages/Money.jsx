import { useState } from 'react';
import { useIdentity } from '../identity';
import { usePoll } from '../usePoll';
import { api } from '../api';
import { nameOf, countdownLabel } from '../utils';

export default function Money({ people, settings }) {
  const identity = useIdentity();
  const payments = usePoll('/payments', 8000);
  const summary = usePoll('/money-summary', 8000);
  const expenses = usePoll('/expenses', 8000);
  const settleUp = usePoll('/settle-up', 8000);
  const [sub, setSub] = useState('payments');

  const peopleList = people.data || [];
  const paymentList = payments.data || [];
  const money = summary.data;
  const pct = money && money.totalNeeded ? Math.min(100, Math.round((money.collected / money.totalNeeded) * 100)) : 0;
  const s = settings.data || {};
  const me = peopleList.find(p => p.id === identity.personId);
  const organizerNames = peopleList.filter(p => p.isOrganizer).map(p => p.name).join(' & ');

  // installments derived from the actual payment rows so they always match what setup created
  const installments = [];
  for (const r of paymentList) {
    if (!installments.find(i => i.dueLabel === r.dueLabel && i.dueDate === r.dueDate)) {
      installments.push({ dueLabel: r.dueLabel, dueDate: r.dueDate, amount: r.amount });
    }
  }

  async function markStatus(p, status) {
    try {
      await api.patch(`/payments/${p.id}`, { status });
      payments.reload();
      summary.reload();
    } catch (e) { alert(e.message); }
  }

  return (
    <div>
      <div className="card">
        <h3>💵 {s.costPerPerson ? `$${s.costPerPerson}/person` : 'Trip cost'}{s.lodgingCost && s.foodCost ? ` — $${s.lodgingCost} lodging + $${s.foodCost} food` : ''}</h3>
        <div className="chip-row">
          {installments.map(i => (
            <span key={i.dueLabel + i.dueDate} className="chip">{i.dueLabel}: ${i.amount.toFixed(2)} due {i.dueDate} ({countdownLabel(i.dueDate)})</span>
          ))}
        </div>
        <div className="small-muted">Pay via Zelle or Apple Pay.</div>
        {money && (
          <>
            <div className="progress-outer"><div className="progress-inner" style={{ width: `${pct}%` }} /></div>
            <div className="small-muted">${money.collected.toFixed(2)} of ${money.totalNeeded.toFixed(2)} collected</div>
          </>
        )}
      </div>

      <div className="subnav">
        <button className={sub === 'payments' ? 'active' : ''} onClick={() => setSub('payments')}>Payment Tracker</button>
        <button className={sub === 'expenses' ? 'active' : ''} onClick={() => setSub('expenses')}>Shared Expenses</button>
      </div>

      {sub === 'payments' && (
        <div className="card">
          {peopleList.filter(p => p.status === 'confirmed' || p.status === 'maybe').map(p => {
            const rows = paymentList.filter(pay => pay.personId === p.id);
            if (!rows.length) return null;
            return (
              <div key={p.id} style={{ marginBottom: 12 }}>
                <strong>{p.name}{p.status === 'maybe' ? ' (maybe)' : ''}</strong>
                {rows.map(r => (
                  <div key={r.id} className="list-item">
                    <span>{r.dueLabel} — ${r.amount.toFixed(2)}</span>
                    <span className="row">
                      <span className={`pill ${r.status === 'confirmed' ? 'organizer' : r.status === 'sent' ? 'goh' : ''}`}>{r.status.replace('_', ' ')}</span>
                      {r.personId === identity.personId && r.status === 'not_sent' && (
                        <button className="btn small ghost" onClick={() => markStatus(r, 'sent')}>Mark Sent</button>
                      )}
                      {me?.isOrganizer && r.status !== 'confirmed' && (
                        <button className="btn small" onClick={() => markStatus(r, 'confirmed')}>Confirm</button>
                      )}
                    </span>
                  </div>
                ))}
              </div>
            );
          })}
          <p className="small-muted">Only organizers{organizerNames ? ` (${organizerNames})` : ''} can mark payments Confirmed.</p>
        </div>
      )}

      {sub === 'expenses' && <ExpensesPanel expenses={expenses} settleUp={settleUp} people={peopleList} identity={identity} />}
    </div>
  );
}

function ExpensesPanel({ expenses, settleUp, people, identity }) {
  const list = expenses.data || [];
  const [description, setDescription] = useState('');
  const [amount, setAmount] = useState('');

  async function add() {
    if (!description.trim() || !amount) return;
    await api.post('/expenses', { description, amount: parseFloat(amount), paidBy: identity.personId, date: new Date().toISOString().slice(0, 10) });
    setDescription(''); setAmount('');
    expenses.reload();
    settleUp.reload();
  }
  async function del(id) {
    await api.del(`/expenses/${id}`);
    expenses.reload();
    settleUp.reload();
  }

  const su = settleUp.data;

  return (
    <div>
      <div className="card">
        <h3>Shared Expenses Ledger</h3>
        {list.map(e => (
          <div key={e.id} className="list-item">
            <span>{e.description} — ${e.amount.toFixed(2)} <span className="small-muted">(paid by {nameOf(people, e.paidBy)}, {e.date})</span></span>
            <button className="btn small ghost" onClick={() => del(e.id)}>×</button>
          </div>
        ))}
        <div className="row" style={{ marginTop: 10 }}>
          <input type="text" placeholder="Description" value={description} onChange={e => setDescription(e.target.value)} />
          <input type="number" placeholder="$" value={amount} onChange={e => setAmount(e.target.value)} style={{ width: 90 }} />
        </div>
        <button className="btn small" style={{ marginTop: 8 }} onClick={add}>Log expense</button>
      </div>

      {su && (
        <div className="card">
          <h3>Settle Up</h3>
          <div className="small-muted">Equal split among confirmed attendees — total ${su.total.toFixed(2)}, ${su.share.toFixed(2)} each</div>
          {su.transactions.length === 0 && <p className="small-muted">Everyone's settled up!</p>}
          {su.transactions.map((t, i) => (
            <div key={i} className="list-item"><span>{t.from} → {t.to}</span><strong>${t.amount.toFixed(2)}</strong></div>
          ))}
        </div>
      )}
    </div>
  );
}
