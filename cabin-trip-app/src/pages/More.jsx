import { useState } from 'react';
import Committees from './Committees.jsx';
import Games from './Games.jsx';
import Logistics from './Logistics.jsx';
import Settings from './Settings.jsx';

export default function More({ people, settings }) {
  const [sub, setSub] = useState('committees');
  return (
    <div>
      <div className="subnav">
        <button className={sub === 'committees' ? 'active' : ''} onClick={() => setSub('committees')}>🎯 Committees</button>
        <button className={sub === 'games' ? 'active' : ''} onClick={() => setSub('games')}>🎲 Games</button>
        <button className={sub === 'logistics' ? 'active' : ''} onClick={() => setSub('logistics')}>🧭 Logistics</button>
        <button className={sub === 'settings' ? 'active' : ''} onClick={() => setSub('settings')}>⚙️ Settings</button>
      </div>
      {sub === 'committees' && <Committees people={people} />}
      {sub === 'games' && <Games people={people} />}
      {sub === 'logistics' && <Logistics people={people} settings={settings} />}
      {sub === 'settings' && <Settings people={people} settings={settings} />}
    </div>
  );
}
