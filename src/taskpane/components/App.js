import * as React from 'react';

export default class App extends React.Component {
  click = async () => {
    return Word.run(async context => {
      /**
       * Insert your Word code here
       */

      // insert a paragraph at the end of the document.
      const paragraph = context.document.body.insertParagraph('Hello World', Word.InsertLocation.end);

      // change the paragraph color to blue.
      paragraph.font.color = 'green';

      await context.sync();
    });
  };

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <p>{title}: Please sideload me!</p>
      );
    }

    return (
      <div>
          <button onClick={this.click}>Try clicking me!</button>
      </div>
    );
  }
}
