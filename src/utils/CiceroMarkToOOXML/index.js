import attachVariableChangeListener from '../AttachVariableChangeListener';

const insertHeading = async (context, value, level) => {
  const definedLevels = {
    1: { style: Word.Style.heading1, size: 25 },
    2: { style: Word.Style.heading2, size: 20 },
    3: { style: Word.Style.heading3, size: 16 },
    4: { style: Word.Style.heading4, size: 16 },
    5: { style: Word.Style.heading5, size: 16 },
    6: { style: Word.Style.heading6, size: 16 },
  };
  const heading = context.document.body.insertParagraph(value, Word.InsertLocation.end);
  heading.styleBuiltIn = definedLevels[level].style;
  heading.font.set({
    highlightColor: null,
    size: definedLevels[level].size,
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
    color: 'black',
    highlightColor: null,
    size: 12,
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
    highlightColor: 'lime',
    size: 12,
  });
  await context.sync();

  // If the app ever goes into an infinite loop, it is probably because of this function call.
  attachVariableChangeListener(context, title);
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

const renderNodes = (context, node, counter, parent=null) => {
  if (node.$class === definedNodes.variable) {
    const tag = node.name;
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
      renderNodes(context, subNode, counter, { class: node.$class, level: node.level });
    });
  }
  if (node.$class === definedNodes.paragraph) {
    node.nodes.forEach(subNode => {
      renderNodes(context, subNode, counter);
    });
  }
};

export default renderNodes;
