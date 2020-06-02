import React, { useState, useEffect } from 'react';

import {
  Library as TemplateLibraryComponent,
} from '@accordproject/cicero-ui';
import { TemplateLibrary, Template, Clause } from '@accordproject/cicero-core';

const LibraryComponent = (props) => {

  const [templates, setTemplates] = useState(null);
  const [templateIndex, setTemplateIndex] = useState(null);
  useEffect(() => {
    async function load() {
      const templateLibrary = new TemplateLibrary();
      const templateIndex = await templateLibrary
        .getTemplateIndex({
          ciceroVersion,
        });
      setTemplateIndex(templateIndex);
      setTemplates(Object.values(templateIndex));
    }
    load();
  },[]);

  if(!templates){
    return (<div>Hey</div>);
  }

  return (
    <TemplateLibraryComponent
      templates = {templates}
      addToCont = { (templateUri) => addToContract(templateIndex, templateUri)}
    />
  );
};

export default LibraryComponent;
