import { CiceroMarkTransformer } from '@accordproject/markdown-cicero';
import { Clause } from '@accordproject/cicero-core';
import { OoxmlTransformer } from '@accordproject/markdown-docx';

const triggerClauseParse = (title, template) => {
  Office.context.document.bindings.addFromNamedItemAsync(title, Office.CoercionType.Text, { id: title }, res => {
    if (res.status === Office.AsyncResultStatus.Succeeded) {
      res.value.addHandlerAsync(Office.EventType.BindingDataChanged, event => textChangeListener(event, template), res => {
        if (res.status === Office.AsyncResultStatus.Succeeded) {
          // ToDo: show the success to user in Production environment
          console.info(`Listener attached to ${title}`);
          return;
        }
        else {
          triggerClauseParse(title);
        }
      });
    }
    else {
      triggerClauseParse(title);
    }
  });
};

const textChangeListener = (event, template) => {
  const { binding } = event;
  binding.getDataAsync({ coercionType: Office.CoercionType.Ooxml }, result => {
    // The OOXML of the clause
    const data = result.value;
    const ooxmlTransformer = new OoxmlTransformer();
    const ciceroMark = ooxmlTransformer.toCiceroMark(data);
    const ciceroMarkTransformer = new CiceroMarkTransformer();
    const inputWrapped = ciceroMarkTransformer.toCiceroMarkUnwrapped(ciceroMark);
    const markdown = ciceroMarkTransformer.toMarkdownCicero(inputWrapped);
    const clause = new Clause(template);
    try {
      clause.parse(markdown);
    }
    catch (error) {
      console.error(error.message);
    }
  });
};

export default triggerClauseParse;
