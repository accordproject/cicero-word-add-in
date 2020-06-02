import React, { useState } from 'react';
import { Menu } from 'semantic-ui-react';

const App = ({ isOfficeInitialized }) => {
  const [activeNav, setActiveNav] = useState('library');

  const handleClick = (event, { name }) => {
    setActiveNav(name);
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
    </div>
  );
};

export default App;
