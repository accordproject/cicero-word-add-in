import React, { useState, useEffect } from 'react';
import { Loader } from 'semantic-ui-react';

import { Library as TemplateLibraryRenderer } from '@accordproject/ui-components';
import { TemplateLibrary, Template, Clause } from '@accordproject/cicero-core';

import renderNodes from '../../utils/CiceroMarkToOOXML';

const XML_HEADER = '<?xml version="1.0" encoding="utf-8" ?>';
const CUSTOM_XML_NAMESPACE = 'https://accordproject.org/';

const LibraryComponent = () => {
  const [templates, setTemplates] = useState(null);
  const [overallCounter, setOverallCounter] = useState({});
  const [selectedTemplates, selectTemplate] = useState({});

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
      Office.context.ui.displayDialogAsync(`${window.location.origin}/bad-file.html`, { width: 30, height: 8 });
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
    const templateIdentifier = `${templateIndex.name}@${templateIndex.version}`;
    selectTemplate({
      ...selectedTemplates,
      [templateIdentifier]: template,
    });
    saveTemplateToXml(templateIdentifier);
  };

  const saveTemplateToXml = templateIdentifier => {
    Office.context.document.customXmlParts.getByNamespaceAsync(CUSTOM_XML_NAMESPACE, result => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        if (result.value.length === 0) {
          const xml = XML_HEADER +
          `<templates xmlns="${CUSTOM_XML_NAMESPACE}">` +
            `<template xmlns="${templateIdentifier}" />` +
          '</templates>';
          Office.context.document.customXmlParts.addAsync(xml);
        }
        else {
          const customXml = result.value[0];
          customXml.getNodesAsync('*/*', result => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              let newXml = XML_HEADER + `<templates xmlns="${CUSTOM_XML_NAMESPACE}">`;
              if (result.value.length > 0) {
                for (let node=0; node < result.value.length; ++node) {
                  if (result.value[node].namespaceUri !== templateIdentifier) {
                    newXml += `<template xmlns="${result.value[node].namespaceUri}" />`;
                  }
                }
              }
              newXml += `<template xmlns="${templateIdentifier}" />`;
              newXml += '</templates>';
              Office.context.document.customXmlParts.getByNamespaceAsync(CUSTOM_XML_NAMESPACE, res => {
                if (res.status === Office.AsyncResultStatus.Succeeded) {
                  for (let index=0; index<res.value.length; ++index) {
                    res.value[index].deleteAsync();
                  }
                }
              });
              Office.context.document.customXmlParts.addAsync(newXml);
            }
          });
        }
      }
    });
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
