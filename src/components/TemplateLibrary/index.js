import React, { useState, useEffect } from 'react';
import { Loader } from 'semantic-ui-react';

import { Library as TemplateLibraryRenderer } from '@accordproject/ui-components';
import { TemplateLibrary, Template, Clause } from '@accordproject/cicero-core';

const LibraryComponent = () => {
  const [templates, setTemplates] = useState(null);

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

  const setup = async (text, data) => {
    await Word.run(async context => {
      // Gets the body of the document
      const body = context.document.body;
      // Inserts sample text at start
      // Ref: https://docs.microsoft.com/en-us/javascript/api/word/word.body?view=word-js-preview#inserttext-text--insertlocation-
      const contractTextRange = body.insertText(text, Word.InsertLocation.start);
      // Search for variables from sample text
      // Ref: https://docs.microsoft.com/en-us/javascript/api/word/word.range?view=word-js-preview#search-searchtext--searchoptions-
      const searchResults = contractTextRange.search('"Party A"');
      // One just needs to do it.
      // Ref: https://youtu.be/22P43aerrho?t=511
      searchResults.load('items/length');
      await context.sync();

      // Inserts content controls
      for(let res=0; res<searchResults.items.length; ++res) {
        // Insert content control where ever "Party A" is found
        // Ref: https://docs.microsoft.com/en-us/javascript/api/word/word.range?view=word-js-preview#insertcontentcontrol--
        let contentControl = searchResults.items[res].insertContentControl();
        contentControl.tag = 'ship';
        contentControl.title = 'Shipper';
        const ooxml = contentControl.getOoxml();
        await context.sync().then(() => {
          if(res==0) {
            console.log(ooxml.value);
          }
        });
      }
      await context.sync();
    });
  };

  const loadTemplateText = async url => {
    // URL to compiled archive
    const template = await Template.fromUrl('https://compiled--templates-accordproject.netlify.app/archives/acceptance-of-delivery@0.13.2.js.cta');

    const sampleText = template.getMetadata().getSample();
    const clause = new Clause(template);
    clause.parse(sampleText);
    const sampleData = clause.getData();
    setup(sampleText, sampleData);
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
