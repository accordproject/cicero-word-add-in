import sanitizeHtmlChars from './SanitizeHtmlChars';

let globalOoxml;

const insertHeading = (value, level) => {
  const definedLevels = {
    1: { style: Word.Style.heading1, size: 25 },
    2: { style: Word.Style.heading2, size: 20 },
    3: { style: Word.Style.heading3, size: 16 },
    4: { style: Word.Style.heading4, size: 16 },
    5: { style: Word.Style.heading5, size: 16 },
    6: { style: Word.Style.heading6, size: 16 },
  };

  return `
    <w:pPr>
      <w:pStyle w:val="${definedLevels[level].style}"/>
    </w:pPr>
    <w:r>
      <w:rPr>
        <w:sz w:val="${definedLevels[level].size * 2}"/>
      </w:rPr>
      <w:t xml:space="preserve">${sanitizeHtmlChars(value)}</w:t>
    </w:r>
  `;
};

const insertLineBreak = () => {
  return '<w:p />';
};

const insertSoftBreak = () => {
  return `
    <w:r>
      <w:sym w:font="Calibri" w:char="2009" />
    </w:r>
  `;
};

const insertText = (value, emphasize=false) => {
  if (emphasize) {
    return `
      <w:r>
        <w:rPr>
          <w:i w:val="true" />
        </w:rPr>
        <w:t>${sanitizeHtmlChars(value)}</w:t>
      </w:r>
    `;
  }
  return `
    <w:r>
      <w:t xml:space="preserve">${sanitizeHtmlChars(value)}</w:t>
    </w:r>
  `;
};

const insertVariable = ( title, tag, value) => {
  return `
    <w:sdt>
      <w:sdtPr>
        <w:rPr>
          <w:color w:val="000000"/>
          <w:sz w:val="24"/>
          <w:highlight w:val="green"/>
        </w:rPr>
        <w:alias w:val="${title}"/>
        <w:tag w:val="${tag}"/>
      </w:sdtPr>
      <w:sdtContent>
        <w:r>
          <w:rPr>
            <w:color w:val="000000"/>
            <w:sz w:val="24"/>
            <w:highlight w:val="green"/>
          </w:rPr>
          <w:t xml:space="preserve">${sanitizeHtmlChars(value)}</w:t>
        </w:r>
      </w:sdtContent>
    </w:sdt>
  `;
};

const insertList = (node, type) => {
  let ooxml = '';
  node.nodes.forEach(subNode => {
    ooxml += `
      <w:p>
        <w:pPr>
          <w:pStyle w:val="ListParagraph"/>
          <w:numPr>
            <w:ilvl w:val="1"/>
            <w:numId w:val="${type === 'ordered' ? 1 : 2}"/>
          </w:numPr>
        </w:pPr>
        ${getListItem(subNode)}
      </w:p>
    `;
  });
  ooxml += insertLineBreak(); // otherwise the last item isn't included in the list
  return ooxml;
};

const getListItem = (node, text='') => {
  if (node.$class === definedNodes.text) {
    return `
      <w:r>
        <w:t xml:space="preserve">${sanitizeHtmlChars(node.text)}</w:t>
      </w:r>
    `;
  }
  if (node.$class === definedNodes.variable) {
    return `
      <w:sdt>
        <w:sdtPr>
          <w:alias w:val="${node.name.toUpperCase()[0]}${node.name.substring(1)}"/>
          <w:tag w:val="${node.name}"/>
        </w:sdtPr>
        <w:sdtContent>
          <w:r>
            <w:rPr>
              <w:highlight w:val="green"/>
            </w:rPr>
            <w:t xml:space="preserve">${sanitizeHtmlChars(node.value)}</w:t>
          </w:r>
        </w:sdtContent>
     </w:sdt>
    `;
  }
  if (node.$class === definedNodes.softbreak) {
    return ' ';
  }
  if (node.nodes !== undefined) {
    node.nodes.forEach(subNode => {
      text += getListItem(subNode, text);
    });
  }
  return text;
};

const definedNodes = {
  computedVariable: 'org.accordproject.ciceromark.ComputedVariable',
  heading: 'org.accordproject.commonmark.Heading',
  item: 'org.accordproject.commonmark.Item',
  list: 'org.accordproject.commonmark.List',
  listBlock: 'org.accordproject.ciceromark.ListBlock',
  paragraph: 'org.accordproject.commonmark.Paragraph',
  softbreak: 'org.accordproject.commonmark.Softbreak',
  text: 'org.accordproject.commonmark.Text',
  variable: 'org.accordproject.ciceromark.Variable',
  emphasize: 'org.accordproject.commonmark.Emph',
};

const getNodes = (node, counter, parent=null) => {
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
    return insertVariable(title, tag, value);
  }
  if (node.$class === definedNodes.text) {
    if (parent !== null && parent.class === definedNodes.heading) {
      return insertHeading(node.text, parent.level);
    }
    else if (parent !== null && parent.class === definedNodes.emphasize) {
      return insertText(node.text, true);
    }
    else {
      return insertText(node.text);
    }
  }
  if (node.$class === definedNodes.softbreak) {
    return insertSoftBreak();
  }
  if (node.$class === definedNodes.emphasize) {
    let ooxml = '';
    node.nodes.forEach(subNode => {
      ooxml += getNodes(subNode, counter, { class: node.$class });
    });
    return ooxml;
  }
  if (node.$class === definedNodes.listBlock || node.$class === definedNodes.list) {
    switch (node.type) {
    case 'ordered':
      globalOoxml += insertList(node, 'ordered');
      break;
    case 'bullet':
      globalOoxml += insertList(node, 'unordered');
      break;
    default:
      globalOoxml;
    }
  }
  if (node.$class === definedNodes.heading) {
    let ooxml = '';
    node.nodes.forEach(subNode => {
      ooxml += getNodes(subNode, counter, { class: node.$class, level: node.level });
    });
    globalOoxml = `
      ${globalOoxml}
      <w:p>
        ${ooxml}
      </w:p>
    `;
  }
  if (node.$class === definedNodes.paragraph) {
    let ooxml = '';
    node.nodes.forEach(subNode => {
      ooxml += getNodes(subNode, counter);
    });
    globalOoxml = `
      ${globalOoxml}
      <w:p>
        ${ooxml}
      </w:p>
    `;
  }
  return '';
};

const ooxmlGenerator = (ciceroMark, counter, ooxml) => {
  globalOoxml = ooxml;
  ciceroMark.nodes.forEach(node => {
    getNodes(node, counter);
  });
  return globalOoxml;
};

export default ooxmlGenerator;
