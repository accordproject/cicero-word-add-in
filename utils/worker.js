import { TemplateLoader } from '@accordproject/cicero-core';

// Receive URL of the template and download it.
onmessage = event => {
  const { url } = event.data;
  const template = TemplateLoader.fromUrl(url);
  self.postMessage(template);
};
