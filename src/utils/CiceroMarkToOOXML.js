import sanitizeHtmlChars from './SanitizeHtmlChars';
import sanitizeStringValues from './SanitizeStringValues';
import titleGenerator from './TitleGenerator';

let globalOoxml;
/**
 * Transforms the given heading node into OOXML heading.
 *
 * @param {string} value Text to be rendered as heading
 * @param {number} level Level of heading - ranges from 1 to 6
 * @returns {string} OOXMl for heading
 *
 */
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

/**
 * Inserts a line break.
 *
 * @returns {string} OOXML for linebreak
 */
const insertLineBreak = () => {
  return '<w:p />';
};

/**
 * Inserts a soft break.
 *
 * @returns {string} OOXML for softbreak
 */
const insertSoftBreak = () => {
  return `
    <w:r>
      <w:sym w:font="Calibri" w:char="2009" />
    </w:r>
  `;
};

/**
 * Inserts text.
 *
 * @param {string}  value     Text to be rendered
 * @param {boolean} emphasize true=emphasized text, false=normal text
 * @returns {string} OOXML for the text
 */
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

/**
 * Inserts a variable.
 *
 * @param {string} title Title of the variable. Eg. receiver-1, shipper-1
 * @param {string} tag   Name of the variable. Eg. receiver, shipper
 * @param {string} value Value of the variable
 * @param {string} type  Type of the variable - Long, Double, etc.
 * @returns {string} OOXML string for the variable
 */
const insertVariable = (title, tag, value, type) => {
  return `
    <w:sdt>
      <w:sdtPr>
        <w:rPr>
          <w:color w:val="000000"/>
          <w:sz w:val="24"/>
          <w:highlight w:val="green"/>
        </w:rPr>
        <w:alias w:val="${titleGenerator(title, type)}"/>
        <w:tag w:val="${tag}"/>
      </w:sdtPr>
      <w:sdtContent>
        <w:r>
          <w:rPr>
            <w:color w:val="000000"/>
            <w:sz w:val="24"/>
            <w:highlight w:val="green"/>
          </w:rPr>
          <w:t xml:space="preserve">${sanitizeStringValues(sanitizeHtmlChars(value))}</w:t>
        </w:r>
      </w:sdtContent>
    </w:sdt>
  `;
};

/**
 * Inserts a list.
 *
 * @param {Array}  node Array of nodes
 * @param {string} type Type of list - ordered or unordered
 * @returns {string} OOXML for list
 */
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

/**
 * Gets the particular list item.
 *
 * @param {Array}  node Array of nodes
 * @param {string} text Text to be rendered
 * @returns {string} OOXML for the list item
 */
const getListItem = (node, text='') => {
  if (node.$class === definedNodes.text) {
    return `
      <w:r>
        <w:t xml:space="preserve">${sanitizeHtmlChars(node.text)}</w:t>
      </w:r>
    `;
  }
  if (node.$class === definedNodes.variable) {
    const name = node.name;
    const type = node.elementType;
    return `
      <w:sdt>
        <w:sdtPr>
          <w:alias w:val="${titleGenerator(name.toUpperCase()[0]+name.substring(1), type)}"/>
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

/**
 * Gets the OOXML for the given node.
 *
 * @param {object} node    Description of node type
 * @param {object} counter Counter for different variables based on node name
 * @param {object} parent  Parent object for a node
 * @returns {string} OOXML for the given node
 */
const getNodes = (node, counter, parent=null) => {
  if (node.$class === definedNodes.variable) {
    const tag = node.name;
    const type = node.elementType;
    if (Object.prototype.hasOwnProperty.call(counter, tag)) {
      counter = {
        ...counter,
        [tag]: {
          ...counter[tag],
          count: ++counter[tag].count,
        },
      };
    }
    else {
      counter[tag] = {
        count: 1,
        type,
      };
    }
    const value = node.value;
    const title = `${tag.toUpperCase()[0]}${tag.substring(1)}${counter[tag].count}`;
    return insertVariable(title, tag, value, type);
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

/**
 * Generates OOXML from CiceroMark JSON.
 *
 * @param {Array}  ciceroMark Ciceromark JSON
 * @param {object} counter    Counter for different variables based on node name
 * @param {string} ooxml      Intial OOXML string
 * @returns {string} Converted OOXML string i.e. CicecoMark->OOXML
 */
const ooxmlGenerator = (ciceroMark, counter, ooxml) => {
  globalOoxml = ooxml;
  ciceroMark.nodes.forEach(node => {
    getNodes(node, counter);
  });
  return globalOoxml;
};

export default ooxmlGenerator;
