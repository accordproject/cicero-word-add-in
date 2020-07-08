const attachVariableChangeListener = (context, title) => {
  Office.context.document.bindings.addFromNamedItemAsync(title, Office.CoercionType.Text, { id: title }, res => {
    if (res.status === Office.AsyncResultStatus.Succeeded) {
      res.value.addHandlerAsync(Office.EventType.BindingDataChanged, variableChangeListener, res => {
        if (res.status === Office.AsyncResultStatus.Succeeded) {
          // ToDo: show the success to user in Production environment
          console.info(`Listener attached to ${title}`);
          return;
        }
        else {
          attachVariableChangeListener(context, title);
        }
      });
    }
    else {
      attachVariableChangeListener(context, title);
    }
  });
};

const variableChangeListener = event => {
  const { binding } = event;
  // ID of the binding the user changed
  const bindingId = binding.id;
  binding.getDataAsync(result => {
    // The text typed by user to change it
    const data = result.value;
    Word.run(async context => {
      // Get all the simlar variables
      const contentControl = context.document.contentControls.getByTitle(bindingId).getFirst();
      contentControl.load('tag');
      await context.sync();
      const tag = contentControl.tag;
      const contentControls = context.document.contentControls.getByTag(tag);
      contentControls.load('items/length');
      await context.sync();
      // To prevent it from an infinite loop, we check if text inside all the variables is same or not.
      let contentControlText = [];
      for(let index=0; index<contentControls.items.length; ++index) {
        let textRange = contentControls.items[index].getRange();
        textRange.load('text');
        await context.sync();
        contentControlText = [textRange.text, ...contentControlText];
      }
      if (contentControlText.every(el => el === contentControlText[0])) {
        return;
      }
      for(let index=0; index<contentControls.items.length; ++index) {
        contentControls.items[index].insertText(data, Word.InsertLocation.replace);
      }
      return context.sync();
    });
  });
};

export default attachVariableChangeListener;
