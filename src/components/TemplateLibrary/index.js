import React, { useState, useEffect, useRef } from 'react';
import { Loader, Divider, Button, Icon } from 'semantic-ui-react';

import { Library as TemplateLibraryRenderer } from '@accordproject/ui-components';
import { TemplateLibrary, Template, Clause } from '@accordproject/cicero-core';

import renderNodes from '../../utils/CiceroMarkToOOXML';

import './index.css';

const LibraryComponent = () => {
  const [templates, setTemplates] = useState(null);
  const [overallCounter, setOverallCounter] = useState({});

  const fileInputRef = useRef(null);

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

  const uploadTemplate = async event => {
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
    <div className="template-container">
      <div className="upload-template">
        <Button icon labelPosition="left" onClick={() => fileInputRef.current.click()}>
          <Icon name="upload" />
          Upload your tempate
          <input
            ref={fileInputRef}
            type="file"
            hidden
            onClick={event => {
              event.persist();
              event.target.value = null;
            }}
            onChange={uploadTemplate}
          />
        </Button>
      </div>
      <Divider horizontal>Or</Divider>
      <div>
        <TemplateLibraryRenderer
          items = {templates}
          onPrimaryButtonClick={loadTemplateText}
          onSecondaryButtonClick={goToTemplateDetail}
        />
      </div>
    </div>
  );
};

export default LibraryComponent;
