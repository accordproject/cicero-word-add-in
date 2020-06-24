const insertHeading = async (context, value, level) => {
  const definedLevels = {
    1: 28,
    2: 26,
    3: 24,
    4: 22,
    5: 20,
    6: 18,
  };
  context.document.body.insertParagraph(value, Word.InsertLocation.end).font.set({
    color: 'black',
    bold: false,
    size: definedLevels[level],
  });
  insertLineBreak(context);
  await context.sync();
};

const insertLineBreak = async context => {
  context.document.body.insertBreak(Word.BreakType.line, Word.InsertLocation.end);
  await context.sync();
};

const insertSoftBreak = async context => {
  context.document.body.insertText(' ', Word.InsertLocation.end);
  await context.sync();
};

const insertText = async (context, value) => {
  context.document.body.insertText(value, Word.InsertLocation.end).font.set({
    bold: false,
    color: 'black',
    size: 14,
  });
  await context.sync();
};

const insertVariable = async (context, title, tag, value) => {
  let variableText = context.document.body.insertText(value, Word.InsertLocation.end);
  let contentControl = variableText.insertContentControl();
  contentControl.title = title;
  contentControl.tag = tag;
  contentControl.font.set({
    color: 'black',
    bold: true,
    size: 14,
  });
  await context.sync();

  setTimeout(() => {
    Office.context.document.bindings.addFromNamedItemAsync(title, 'text', { id: title }, res => {
      if (res.status === Office.AsyncResultStatus.Succeeded) {
        res.value.addHandlerAsync(Office.EventType.BindingDataChanged, handler);
      }
      else {
        // ToDo: show the error to user in Production environment
        console.error(title, res);
      }
    });
  }, 100);
  await context.sync();
};

const definedNodes = {
  computedVariable: 'org.accordproject.ciceromark.ComputedVariable',
  heading: 'org.accordproject.commonmark.Heading',
  item: 'org.accordproject.commonmark.Item',
  list: 'org.accordproject.commonmark.List',
  listVariable: 'org.accordproject.ciceromark.ListVariable',
  paragraph: 'org.accordproject.commonmark.Paragraph',
  softbreak: 'org.accordproject.commonmark.Softbreak',
  text: 'org.accordproject.commonmark.Text',
  variable: 'org.accordproject.ciceromark.Variable',
};

let counter = {};

const handler = event => {
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

const renderNodes = (context, node, parent=null) => {
  if (node.$class === definedNodes.variable) {
    const tag = node.id;
    if (Object.prototype.hasOwnProperty.call(counter, tag)) {
      counter = {
        ...counter,
        [tag]: ++counter[tag],
      };
    }
    else {
      counter[tag] = 1;
    }
    const value = node.value;
    const title = `${tag.toUpperCase()[0]}${tag.substring(1)}${counter[tag]}`;
    insertVariable(context, title, tag, value);
    return;
  }
  if (node.$class === definedNodes.text) {
    if (parent !== null && parent.class === definedNodes.heading) {
      insertHeading(context, node.text, parent.level);
    }
    else {
      insertText(context, node.text);
    }
    return;
  }
  if (node.$class === definedNodes.softbreak) {
    insertSoftBreak(context);
    return;
  }
  if (node.$class === definedNodes.heading) {
    node.nodes.forEach(subNode => {
      renderNodes(context, subNode, { class: node.$class, level: node.level });
    });
  }
  if (node.$class === definedNodes.paragraph) {
    node.nodes.forEach(subNode => {
      renderNodes(context, subNode);
    });
  }
};

export default renderNodes;
