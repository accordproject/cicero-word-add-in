import React, { useState, useEffect } from 'react';
import { Loader } from 'semantic-ui-react';

import { Library as TemplateLibraryRenderer } from '@accordproject/cicero-ui';
import { TemplateLibrary } from '@accordproject/cicero-core';

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

  if(!templates){
    return <Loader active>Loading</Loader>;
  }

  return (
    <TemplateLibraryRenderer
      items = {templates}
      // TODO
      onPrimaryButtonClick={() => console.log('Action to add this template to contract')}
      onSecondaryButtonClick={() => console.log('Action to view the details')}
    />
  );
};

export default LibraryComponent;
