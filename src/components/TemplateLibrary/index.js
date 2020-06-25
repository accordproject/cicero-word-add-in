import React, { useState, useEffect } from 'react';
import { Loader } from 'semantic-ui-react';

import { Library as TemplateLibraryRenderer } from '@accordproject/ui-components';
import { TemplateLibrary, Template, Clause } from '@accordproject/cicero-core';
import { CiceroMarkTransformer } from '@accordproject/markdown-cicero';

import renderNodes from '../../utils/CiceroMarkToOOXML';
import variableChangeListener from '../../utils/VariableChangeListener';

const LibraryComponent = () => {
  const [templates, setTemplates] = useState(null);
  const [overallCounter, setOverallCounter] = useState({});

  useEffect(() => {
    async function load() {
      const templateLibrary = new TemplateLibrary();
      const templateIndex = await templateLibrary
        .getTemplateIndex({
          latestVersion: true,
        });
      setTemplates(Object.values(templateIndex));
    }
    load();
  }, []);

  const setup = async dom => {
    await Word.run(async context => {
      let counter = {};
      dom.nodes.forEach(node => {
        renderNodes(context, node, counter);
      });
      await context.sync();
      for (const variable in counter) {
        for (let count=1; count<=counter[variable]; ++count) {
          const title = `${variable.toUpperCase()[0]}${variable.substring(1)}${count}`;
          await attachVariableChangeListener(title);
        }
      }
      setOverallCounter({
        ...overallCounter,
        ...counter,
      });
    });
  };

  const attachVariableChangeListener = async title => {
    await Word.run(async context => {
      Office.context.document.bindings.addFromNamedItemAsync(title, 'text', { id: title }, res => {
        if (res.status === Office.AsyncResultStatus.Succeeded) {
          res.value.addHandlerAsync(Office.EventType.BindingDataChanged, variableChangeListener, res => {
            if (res.status === Office.AsyncResultStatus.Succeeded) {
              // ToDo: show the success to user in Production environment
              console.info(`Listener attached to ${title}`);
            }
            else {
              // ToDo: show the error to user in Production environment
              console.error(`Listener failed to attach to ${title}`);
            }
          });
        }
        else {
          // ToDo: show the error to user in Production environment
          console.error(title, res);
        }
      });
      await context.sync();
    });
  };

  const loadTemplateText = async url => {
    // URL to compiled archive
    const template = await Template.fromUrl('https://compiled--templates-accordproject.netlify.app/archives/acceptance-of-delivery@0.13.2.js.cta');

    const sampleText = template.getMetadata().getSample();
    const clause = new Clause(template);
    clause.parse(sampleText);
    const sampleWrapped = await clause.draft({ wrapVariables: true });
    const ciceroMarkTransformer = new CiceroMarkTransformer();
    const ciceroDOM = ciceroMarkTransformer.fromMarkdown(sampleWrapped, 'json');
    setup(ciceroDOM);
  };

  const goToTemplateDetail = template => {
    const templateOrigin = new URL(template.url).origin;
    const { name, version } = template;
    window.open(`${templateOrigin}/${name}@${version}.html`, '_blank');
  };

  if(!templates){
    return <Loader active>Loading</Loader>;
  }

  return (
    <TemplateLibraryRenderer
      items = {templates}
      onPrimaryButtonClick={template => loadTemplateText(template.url)}
      onSecondaryButtonClick={template => goToTemplateDetail(template)}
    />
  );
};

export default LibraryComponent;
