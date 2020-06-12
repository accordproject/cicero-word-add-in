import React, { useState, useEffect } from 'react';
import { Loader } from 'semantic-ui-react';

import { Library as TemplateLibraryRenderer } from '@accordproject/ui-components';
import { TemplateLibrary } from '@accordproject/cicero-core';


const LibraryComponent = () => {
  const [templates, setTemplates] = useState(null);
  const [worker, setWorker] = useState(null);

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

    if (window.Worker) {
      setWorker(new Worker('../../../utils/worker.js', { type: 'module' }));
    }
  }, []);

  useEffect(() => {
    // Receive template from worker
    if (worker) {
      worker.onmessage = event => {
        console.log(event.data);
      };
    }
  }, [worker]);

  const loadTemplateText = async url => {
    // Checks if there is an instance of `Worker` and posts message (URL of the template) to it.
    if (worker) {
      worker.postMessage({ url });
    }
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
