import App from './components/App';
import { AppContainer } from 'react-hot-loader';
import * as React from 'react';
import * as ReactDOM from 'react-dom';

let isOfficeInitialized = false;
/**
 * Renders a react component.
 *
 * @param {React.ReactDOM} Component React component to be rendered
 */
const render = Component => {
  ReactDOM.render(
    <AppContainer>
      <Component isOfficeInitialized={isOfficeInitialized} />
    </AppContainer>,
    document.getElementById('container')
  );
};

/* Renders application after Office initializes */
Office.initialize = () => {
  isOfficeInitialized = true;
  render(App);
};

/* Initial render showing a progress bar */
render(App);

if (module.hot) {
  module.hot.accept('./components/App', () => {
    const NextApp = require('./components/App').default;
    render(NextApp);
  });
}
