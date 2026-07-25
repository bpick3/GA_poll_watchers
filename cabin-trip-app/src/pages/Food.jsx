import { useState } from 'react';
import { useIdentity } from '../identity';
import { usePoll } from '../usePoll';
import { api } from '../api';
import { nameOf, csvToNames } from '../utils';

const MEAL_TYPES = ['Breakfast', 'Lunch', 'Dinner', 'Snacks'];
const CATEGORIES = ['Produce', 'Protein', 'Pantry', 'Drinks', 'Snacks', 'Paper Goods'];

export default function Food({ people }) {
  const identity = useIdentity();
  const [sub, setSub] = useState('meals');
  const days = usePoll('/days', 10000);
  const meals = usePoll('/meals', 8000);
  const dietary = usePoll('/dietary-notes', 15000);
  const groceries = usePoll('/groceries', 8000);
  const drinks = usePoll('/drinks-snacks', 8000);

  const peopleList = people.data || [];
  const dayList = days.data || [];

  return (
    <div>
      <div className="subnav">
        <button className={sub === 'meals' ? 'active' : ''} onClick={() => setSub('meals')}>Meal Grid</button>
        <button className={sub === 'dietary' ? 'active' : ''} onClick={() => setSub('dietary')}>Dietary</button>
        <button className={sub === 'groceries' ? 'active' : ''} onClick={() => setSub('groceries')}>Groceries</button>
        <button className={sub === 'drinks' ? 'active' : ''} onClick={() => setSub('drinks')}>Drinks & Snacks</button>
      </div>

      {sub === 'meals' && <MealGrid days={dayList} meals={meals} people={peopleList} />}
      {sub === 'dietary' && <Dietary dietary={dietary} people={peopleList} identity={identity} />}
      {sub === 'groceries' && <Groceries groceries={groceries} identity={identity} />}
      {sub === 'drinks' && <Drinks drinks={drinks} identity={identity} />}
    </div>
  );
}

function MealGrid({ days, meals, people }) {
  const mealList = meals.data || [];
  if (!days.length) return <p className="small-muted">Loading…</p>;

  async function claim(meal, field) {
    const current = (meal[field] || '').split(',').filter(Boolean);
    const idIdx = current.indexOf(localStorage.getItem('cabin_person_id'));
    let next;
    if (idIdx >= 0) next = current.filter((_, i) => i !== idIdx);
    else next = [...current, localStorage.getItem('cabin_person_id')];
    await api.patch(`/meals/${meal.id}`, { [field]: next.join(',') });
    meals.reload();
  }

  return (
    <div className="card" style={{ overflowX: 'auto' }}>
      <table className="grid">
        <thead>
          <tr>
            <th>Meal</th>
            {days.map(d => <th key={d.id}>{d.label}</th>)}
          </tr>
        </thead>
        <tbody>
          {MEAL_TYPES.map(mt => (
            <tr key={mt}>
              <td><strong>{mt}</strong></td>
              {days.map(d => {
                const meal = mealList.find(m => m.dayId === d.id && m.mealType === mt);
                if (!meal) return <td key={d.id}>—</td>;
                const needsOwner = meal.plan && !meal.cooks;
                return (
                  <td key={d.id} style={needsOwner ? { background: 'rgba(255,184,77,0.25)', borderRadius: 8 } : {}}>
                    <div>{meal.plan || <span className="small-muted">open</span>}</div>
                    <div className="small-muted">
                      Cook: {csvToNames(meal.cooks, people).join(', ') || <span className="needs-owner">need one!</span>}
                    </div>
                    <div className="small-muted">Cleanup: {csvToNames(meal.cleanup, people).join(', ') || '—'}</div>
                    <div className="row" style={{ marginTop: 4 }}>
                      <button className="btn small ghost" onClick={() => claim(meal, 'cooks')}>Cook</button>
                      <button className="btn small ghost" onClick={() => claim(meal, 'cleanup')}>Cleanup</button>
                    </div>
                  </td>
                );
              })}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function Dietary({ dietary, people, identity }) {
  const notes = dietary.data || [];
  const [text, setText] = useState((notes.find(n => n.personId === identity.personId) || {}).note || '');

  async function save() {
    await api.put(`/dietary-notes/${identity.personId}`, { note: text });
    dietary.reload();
  }

  return (
    <div>
      <div className="card">
        <h3>My dietary note</h3>
        <div className="field"><textarea value={text} onChange={e => setText(e.target.value)} placeholder="e.g. Vegetarian, nut allergy…" /></div>
        <button className="btn small" onClick={save}>Save</button>
      </div>
      <div className="card">
        <h3>Everyone's notes</h3>
        {notes.filter(n => n.note).map(n => (
          <div key={n.personId} className="list-item"><span>{nameOf(people, n.personId)}</span><span>{n.note}</span></div>
        ))}
        {notes.filter(n => n.note).length === 0 && <p className="small-muted">No dietary notes yet.</p>}
      </div>
    </div>
  );
}

function Groceries({ groceries, identity }) {
  const list = groceries.data || [];
  const [category, setCategory] = useState(CATEGORIES[0]);
  const [item, setItem] = useState('');
  const [qty, setQty] = useState('');

  // union of the suggested defaults + any custom categories people have typed in,
  // so a custom category still gets its own section instead of disappearing
  const allCategories = [...new Set([...CATEGORIES, ...list.map(g => g.category)])];

  async function add() {
    if (!item.trim()) return;
    await api.post('/groceries', { category: category.trim() || 'Other', item, qty });
    setItem(''); setQty('');
    groceries.reload();
  }
  async function claim(g) {
    const mine = localStorage.getItem('cabin_person_name');
    await api.patch(`/groceries/${g.id}`, { claimedBy: g.claimedBy ? null : mine, checked: g.checked });
    groceries.reload();
  }
  async function check(g) {
    await api.patch(`/groceries/${g.id}`, { claimedBy: g.claimedBy, checked: !g.checked });
    groceries.reload();
  }
  async function remove(id) {
    await api.del(`/groceries/${id}`);
    groceries.reload();
  }

  return (
    <div>
      <div className="card">
        <h3>🛒 Shared Grocery List</h3>
        {allCategories.map(cat => {
          const items = list.filter(g => g.category === cat);
          if (!items.length) return null;
          return (
            <div key={cat} style={{ marginBottom: 10 }}>
              <div className="section-title" style={{ margin: '10px 0 4px' }}><strong>{cat}</strong></div>
              {items.map(g => (
                <div key={g.id} className={`checklist-item ${g.checked ? 'done' : ''}`}>
                  <input type="checkbox" checked={!!g.checked} onChange={() => check(g)} />
                  <span style={{ flex: 1 }}>{g.item} {g.qty && `(${g.qty})`}</span>
                  <button className="btn small ghost" onClick={() => claim(g)}>{g.claimedBy || 'Claim'}</button>
                  <button className="btn small danger" onClick={() => remove(g.id)}>×</button>
                </div>
              ))}
            </div>
          );
        })}
      </div>
      <div className="card">
        <h3>+ Add item</h3>
        <div className="field">
          <label>Category</label>
          <input type="text" list="grocery-categories" value={category} onChange={e => setCategory(e.target.value)} />
          <datalist id="grocery-categories">{allCategories.map(c => <option key={c} value={c} />)}</datalist>
        </div>
        <div className="row">
          <input type="text" placeholder="Item" value={item} onChange={e => setItem(e.target.value)} />
          <input type="text" placeholder="Qty" value={qty} onChange={e => setQty(e.target.value)} style={{ width: 90 }} />
        </div>
        <button className="btn small" style={{ marginTop: 8 }} onClick={add}>Add</button>
      </div>
    </div>
  );
}

function Drinks({ drinks, identity }) {
  const list = drinks.data || [];
  const [item, setItem] = useState('');

  async function claim(d) {
    const mine = localStorage.getItem('cabin_person_name');
    await api.patch(`/drinks-snacks/${d.id}`, { claimedBy: d.claimedBy ? null : mine });
    drinks.reload();
  }
  async function add() {
    if (!item.trim()) return;
    await api.post('/drinks-snacks', { item });
    setItem('');
    drinks.reload();
  }

  return (
    <div className="card">
      <h3>🍻 Drinks & Snacks Board</h3>
      {list.map(d => (
        <div key={d.id} className="list-item">
          <span>{d.item}</span>
          <button className={`btn small ${d.claimedBy ? '' : 'ghost'}`} onClick={() => claim(d)}>{d.claimedBy || 'Claim it'}</button>
        </div>
      ))}
      <div className="row" style={{ marginTop: 10 }}>
        <input type="text" placeholder="Add something you'll bring" value={item} onChange={e => setItem(e.target.value)} />
        <button className="btn small" onClick={add}>Add</button>
      </div>
    </div>
  );
}
