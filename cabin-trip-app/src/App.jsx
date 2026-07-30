import { useState, useEffect } from 'react';
import { IdentityContext, useIdentityState } from './identity';
import { usePoll } from './usePoll';
import { api } from './api';
import RosterPicker from './pages/RosterPicker.jsx';
import Setup from './pages/Setup.jsx';
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
  const [moreSub, setMoreSub] = useState(null);

  function openMore(subTab) {
    setMoreSub(subTab);
    setTab('more');
  }

  const people = usePoll('/people', 10000);
  const settings = usePoll('/settings', 15000);
  const s = settings.data || {};

  if (settings.data && s.setupComplete !== '1') {
    return (
      <IdentityContext.Provider value={identity}>
        <Setup settings={settings} />
      </IdentityContext.Provider>
    );
  }

  if (!settings.data) {
    return <div className="app-loading">🍂 Loading…</div>;
  }

  if (!identity.personId) {
    return (
      <IdentityContext.Provider value={identity}>
        <RosterPicker peopleData={people} settings={settings} />
      </IdentityContext.Provider>
    );
  }

  const gohNames = (people.data || []).filter(p => p.isGuestOfHonor).map(p => `${p.name} 🎂`).join(' · ');
  const dateRangeLabel = s.tripStart && s.tripEnd
    ? `${new Date(s.tripStart + 'T00:00:00').toLocaleDateString('en-US', { month: 'short', day: 'numeric' })}–${new Date(s.tripEnd + 'T00:00:00').toLocaleDateString('en-US', { day: 'numeric' })}`
    : '';

  return (
    <IdentityContext.Provider value={identity}>
      <div className="app-shell">
        <header className="app-header">
          <div>
            <h1>🏔️ {s.tripName || 'Cabin Fever 2026'}</h1>
            <div className="subtitle">{gohNames}{gohNames && dateRangeLabel ? ' — ' : ''}{dateRangeLabel}</div>
          </div>
          <button className="whoami" onClick={() => identity.clearPerson()}>{identity.personName} ⏷</button>
        </header>
        <div className="page">
          {tab === 'home' && <Home people={people} settings={settings} setTab={setTab} openMore={openMore} />}
          {tab === 'schedule' && <Schedule people={people} />}
          {tab === 'food' && <Food people={people} />}
          {tab === 'money' && <Money people={people} settings={settings} />}
          {tab === 'more' && <More people={people} settings={settings} initialSub={moreSub} />}
        </div>
        <nav className="tabs-bottom">
          {TABS.map(t => (
            <button
              key={t.key}
              className={`tab-btn ${tab === t.key ? 'active' : ''}`}
              onClick={() => { if (t.key === 'more') setMoreSub(null); setTab(t.key); }}
            >
              <span className="icon">{t.icon}</span>
              <span>{t.label}</span>
            </button>
          ))}
        </nav>
      </div>
    </IdentityContext.Provider>
  );
}
