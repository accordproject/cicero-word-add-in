import React, { useState, useEffect } from 'react';
import { Loader } from 'semantic-ui-react';

import { Library as TemplateLibraryRenderer } from '@accordproject/ui-components';
import { TemplateLibrary, Template, Clause } from '@accordproject/cicero-core';
import { CiceroMarkTransformer } from '@accordproject/markdown-cicero';

import renderNodes from '../../utils/CiceroMarkToOOXML';

const LibraryComponent = () => {
  const [templates, setTemplates] = useState(null);
  const [overallCounter, setOverallCounter] = useState({});

  useEffect(() => {
    async function load() {
      const templateLibrary = new TemplateLibrary('https://deploy-preview-344--templates-accordproject.netlify.app');
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
      setOverallCounter({
        ...overallCounter,
        ...counter,
      });
    });
  };

  const loadTemplateText = async templateIndex => {
    // URL to compiled archive
    const url = new URL(templateIndex.ciceroUrl);
    const template = await Template.fromUrl(`https://deploy-preview-344--templates-accordproject.netlify.app${url.pathname}`);
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
      onPrimaryButtonClick={loadTemplateText}
      onSecondaryButtonClick={goToTemplateDetail}
    />
  );
};

export default LibraryComponent;
