import titleGenerator from './TitleGenerator';
/**
 * Class to visit the variables using the CiceroMark JSON of a template.
 */
class VariableVisitor {
  /**
   * Visits the children.
   *
   * @param {Function} visitor    Visit the variable
   * @param {object}   thing      CiceroMark object of a template
   * @param {object}   parameters Various functions for an object
   * @param {Array}    result     Result array
   * @param {string}   field      Field name
   */
  static visitChildren(visitor, thing, parameters, result, field = 'nodes') {
    if (thing[field]) {
      VariableVisitor.visitNodes(visitor, thing[field], parameters, result);
    }
  }

  /**
   * Visits the nodes for fields other than Variable type.
   *
   * @param {string} visitor    Type of VariableVisitor
   * @param {object} things     CiceroMark JSON for the field
   * @param {object} parameters Counter for variables
   * @param {Array}  result     Variable titles
   */
  static visitNodes(visitor, things, parameters, result) {
    things.forEach(node => {
      visitor.visit(node, parameters, result);
    });
  }

  /**
   * Updates the counter for variable fields and visits the sub fields for other types.
   *
   * @param {object} thing      CiceroMark JSON of template
   * @param {object} parameters Count of different variables
   * @param {Array}  result     Variable Titles
   */
  static visit(thing, parameters, result) {
    switch (thing.$class) {
    case 'org.accordproject.ciceromark.Variable': {
      const variableName = thing.name;
      if (Object.prototype.hasOwnProperty.call(parameters, variableName)) {
        parameters = {
          ...parameters,
          [variableName]: {
            ...parameters[variableName],
            count: ++parameters[variableName].count,
          },
        };
      }
      else {
        parameters[variableName] = {
          count: 1,
          type: thing.elementType,
        };
      }
      result.push(titleGenerator(`${variableName.toUpperCase()[0]}${variableName.substring(1)}${parameters[variableName].count}`, `${parameters[variableName].type}`));
    }
      break;
    default:
      VariableVisitor.visitChildren(this, thing, parameters, result);
    }
  }

  /**
   * Visits the variables present in CiceroMark JSON and returns array of variable fields.
   *
   * @param {object} input CiceroMark JSON of a template
   * @returns {Array} Variable fields for the JSON
   */
  static getVariables(input) {
    const parameters = {};
    const result = [];
    VariableVisitor.visit(input, parameters, result);
    return result;
  }
}

export default VariableVisitor;
