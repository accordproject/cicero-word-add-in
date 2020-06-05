import React, { useState, useEffect } from 'react';
import { Menu } from 'semantic-ui-react';

import './App.css';

const App = ({ isOfficeInitialized }) => {
  const [activeNav, setActiveNav] = useState('library');
  const [openOnStartup, setOpenOnStartup] = useState(false);

  useEffect(() => {
    if (isOfficeInitialized) {
      const autoOpenSetting = Office.context.document.settings.get('Office.AutoShowTaskpaneWithDocument');
      setOpenOnStartup(autoOpenSetting);
    }
  }, [isOfficeInitialized]);

  const handleClick = (event, { name }) => {
    setActiveNav(name);
  };

  const handleStartupState = (event) => {
    Office.context.document.settings.set('Office.AutoShowTaskpaneWithDocument', event.target.checked);
    setOpenOnStartup(event.target.checked);
    Office.context.document.settings.saveAsync();
  };

  const navItems = [
    { name: 'document', content: 'Document', component: <p>Document component goes here.</p> },
    { name: 'library', content: 'Library', component: <p>Library component goes here.</p> },
  ];

  if (!isOfficeInitialized) {
    return (
      <p>Please sideload the extension.</p>
    );
  }

  return (
    <div>
      <Menu widths={navItems.length}>
        {navItems.map((item, index) => (
          <Menu.Item
            active={activeNav === item.name}
            key={index}
            name={item.name}
            onClick={handleClick}
            content={item.content}
          />
        ))}
      </Menu>
      {navItems.map((item, index) => (
        item.name === activeNav && item.component
      ))}
      <footer className="startup-container">
        <label className="checkbox">
          <span>Auto open on startup:</span>
          <input type="checkbox" checked={openOnStartup} onChange={handleStartupState}></input>
        </label>
      </footer>
    </div>
  );
};

export default App;
