import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { Message } from 'semantic-ui-react';

import './index.css';
/**
 * Renders a message when a bad file is uploaded.
 *
 * @returns {React.FC} Bad file component
 */
const BadFile = () => (
  <Message negative className="message-container">
    <Message.Header>Bad file uploaded</Message.Header>
    <p>The file uploaded is not a valid cicero template.</p>
  </Message>
);

ReactDOM.render(<BadFile />, document.getElementById('bad-file'));
