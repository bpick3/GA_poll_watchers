import { useState, useEffect } from 'react';
import { IdentityContext, useIdentityState } from './identity';
import { usePoll } from './usePoll';
import { api } from './api';
import RosterPicker from './pages/RosterPicker.jsx';
import Home from './pages/Home.jsx';
import Schedule from './pages/Schedule.jsx';
import Food from './pages/Food.jsx';
import Money from './pages/Money.jsx';
import More from './pages/More.jsx';

const TABS = [
  { key: 'home', label: 'Home', icon: '🏔️' },
  { key: 'schedule', label: 'Schedule', icon: '🔥' },
  { key: 'food', label: 'Food', icon: '🌮' },
  { key: 'money', label: 'Money', icon: '💵' },
  { key: 'more', label: 'More', icon: '🍂' },
];

export default function App() {
  const identity = useIdentityState();
  const [tab, setTab] = useState('home');
  const people = usePoll('/people', 10000);
  const settings = usePoll('/settings', 15000);

  if (!identity.personId) {
    return (
      <IdentityContext.Provider value={identity}>
        <RosterPicker peopleData={people} />
      </IdentityContext.Provider>
    );
  }

  return (
    <IdentityContext.Provider value={identity}>
      <div className="app-shell">
        <header className="app-header">
          <div>
            <h1>🏔️ Cabin Trip 2026</h1>
            <div className="subtitle">Brandon 🎂 · Rachel 🎂 · Lance 🎂 — Oct 2–5</div>
          </div>
          <button className="whoami" onClick={() => identity.clearPerson()}>{identity.personName} ⏷</button>
        </header>
        <div className="page">
          {tab === 'home' && <Home people={people} settings={settings} setTab={setTab} />}
          {tab === 'schedule' && <Schedule people={people} />}
          {tab === 'food' && <Food people={people} />}
          {tab === 'money' && <Money people={people} settings={settings} />}
          {tab === 'more' && <More people={people} settings={settings} />}
        </div>
        <nav className="tabs-bottom">
          {TABS.map(t => (
            <button key={t.key} className={`tab-btn ${tab === t.key ? 'active' : ''}`} onClick={() => setTab(t.key)}>
              <span className="icon">{t.icon}</span>
              <span>{t.label}</span>
            </button>
          ))}
        </nav>
      </div>
    </IdentityContext.Provider>
  );
}
