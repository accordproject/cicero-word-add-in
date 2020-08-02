import React, { useState, useEffect, useRef } from 'react';
import { Loader } from 'semantic-ui-react';

import { Library as TemplateLibraryRenderer } from '@accordproject/ui-components';
import { TemplateLibrary, Template, Clause } from '@accordproject/cicero-core';

import renderNodes from '../../utils/CiceroMarkToOOXML';

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

  const onUploadTemplate = async event => {
    const fileUploaded = event.target.files[0];
    try {
      const template = await Template.fromArchive(fileUploaded);
      setup(template);
    }
    catch (error) {
      // show error
    }
  };

  const setup = async template => {
    const sampleText = template.getMetadata().getSample();
    const clause = new Clause(template);
    clause.parse(sampleText);
    const ciceroMark = clause.draft({ format : 'ciceromark_parsed' });
    await Word.run(async context => {
      let counter = {};
      context.document.body.insertBreak(Word.BreakType.line, Word.InsertLocation.end);
      ciceroMark.nodes.forEach(node => {
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
    const template = await Template.fromUrl(templateIndex.ciceroUrl);
    setup(template);
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
      onUploadItem={onUploadTemplate}
    />
  );
};

export default LibraryComponent;
