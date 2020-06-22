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
      // Don't search for these attribiutes in the document
      const attributesToSkip = ['clauseId', '$class'];
      for(const key in data) {
        if (attributesToSkip.includes(key)) {
          continue;
        }
        // Search for variables from sample text
        // Ref: https://docs.microsoft.com/en-us/javascript/api/word/word.range?view=word-js-preview#search-searchtext--searchoptions-
        let searchResults = contractTextRange.search(`"${data[key]}"`);
        // One just needs to do it.
        // Ref: https://youtu.be/22P43aerrho?t=511
        searchResults.load('items/length');
        await context.sync();
        // Inserts content controls
        for(let res=0; res<searchResults.items.length; ++res) {
          // Insert content control where ever "Party A" is found
          // Ref: https://docs.microsoft.com/en-us/javascript/api/word/word.range?view=word-js-preview#insertcontentcontrol--
          let contentControl = searchResults.items[res].insertContentControl();
          contentControl.tag = key.toLocaleLowerCase();
          contentControl.title = `${key[0].toUpperCase()}${key.substring(1)}${res+1}`;
        }
        await context.sync();
        // Add an event listener which shall respond when user changes the value of one of the variables.
        for(let res=0; res<searchResults.items.length; ++res) {
          Office.context.document.bindings.addFromNamedItemAsync(`${key[0].toUpperCase()}${key.substring(1)}${res+1}`, 'text', { id: `${key[0].toUpperCase()}${key.substring(1)}${res+1}` }, (res) => {
            res.value.addHandlerAsync(Office.EventType.BindingDataChanged, handler);
          });
        }
        await context.sync();
      }
    });
  };

  const handler = (event) => {
    const { binding } = event;
    // ID of the binding the user changed
    const bindingId = binding.id;
    binding.getDataAsync((result) => {
      // The text typed by user to change it
      const data = result.value;
      Word.run(async context => {
        // Get all the simlar variables
        const contentControl = context.document.contentControls.getByTitle(bindingId).getFirst();
        contentControl.load('tag');
        await context.sync();
        const tag = contentControl.tag;
        const contentControls = context.document.contentControls.getByTag(tag);
        contentControls.load('items/length');
        await context.sync();
        // To prevent it from an infinite loop, we check if text inside all the variables is same or not.
        let contentControlText = [];
        for(let index=0; index<contentControls.items.length; ++index) {
          let textRange = contentControls.items[index].getRange();
          textRange.load('text');
          await context.sync();
          contentControlText = [textRange.text, ...contentControlText];
        }
        if (contentControlText.every((el) => el === contentControlText[0])) {
          return;
        }
        for(let index=0; index<contentControls.items.length; ++index) {
          contentControls.items[index].insertText(data, Word.InsertLocation.replace);
        }
        return context.sync();
      })
    })
  }

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
