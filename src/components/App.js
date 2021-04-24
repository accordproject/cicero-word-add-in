import React, { useState, useEffect } from 'react';
import PropTypes from 'prop-types';
import { Menu } from 'semantic-ui-react';
import { SemanticToastContainer } from 'react-semantic-toasts';
import 'react-semantic-toasts/styles/react-semantic-alert.css';

import TemplateLibrary from './TemplateLibrary';
import './App.css';

/**
 * Returns the App component which is rendered.
 *
 * @param {boolean} isOfficeInitialized Checks whether office is intialized or not
 * @returns {React.FC} React component
 */
const App = ({ isOfficeInitialized }) => {
  const [activeNav, setActiveNav] = useState('library');
  const [openOnStartup, setOpenOnStartup] = useState(false);

  useEffect(() => {
    if (isOfficeInitialized) {
      const autoOpenSetting = Office.context.document.settings.get('Office.AutoShowTaskpaneWithDocument');
      setOpenOnStartup(autoOpenSetting);
    }
  }, [isOfficeInitialized]);

  /**
   * Sets the active nav item in the navbar.
   *
   * @param {MouseEvent} event    Mouse click event
   * @param {object}     obj      An object
   * @param {string}     obj.name Name of the event
   */
  const handleClick = (event, { name }) => {
    setActiveNav(name);
  };

  /**
   * Sets whether add-in should load by default on MS Word startup.
   *
   * @param {MouseEvent} event Mouseclick to see if checkbox is clicked
   */
  const handleStartupState = event => {
    Office.context.document.settings.set('Office.AutoShowTaskpaneWithDocument', event.target.checked);
    setOpenOnStartup(event.target.checked);
    Office.context.document.settings.saveAsync();
  };

  const navItems = [
    { name: 'document', content: 'Document', component: <p>Document component goes here.</p> },
    { name: 'library', content: 'Library', component: <TemplateLibrary /> },
  ];

  if (!isOfficeInitialized) {
    return (
      <p>Please sideload the extension.</p>
    );
  }

  return (
    <React.Fragment>
      <SemanticToastContainer position="top-center" />
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
      <div className="menu-body">
        {navItems.map(item => (
          item.name === activeNav && item.component
        ))}
      </div>
      <footer className="startup-container">
        <label className="checkbox">
          <span>Auto open on startup:</span>
          <input type="checkbox" checked={openOnStartup} onChange={handleStartupState}></input>
        </label>
      </footer>
    </React.Fragment>
  );
};

App.propTypes = {
  isOfficeInitialized: PropTypes.bool,
};

export default App;
