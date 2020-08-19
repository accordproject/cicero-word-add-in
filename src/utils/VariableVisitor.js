import titleGenerator from "./TitleGenerator";

class VariableVisitor {
  static visitChildren(visitor, thing, parameters, result, field = 'nodes') {
    if(thing[field]) {
      VariableVisitor.visitNodes(visitor, thing[field], parameters, result);
    }
  }

  static visitNodes(visitor, things, parameters, result) {
    things.forEach(node => {
      visitor.visit(node, parameters, result);
    });
  }

  static visit(thing, parameters, result) {
    switch(thing.$class) {
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

  static getVariables(input) {
    const parameters = {};
    const result = [];
    VariableVisitor.visit(input, parameters, result);
    return result;
  }
}

export default VariableVisitor;
