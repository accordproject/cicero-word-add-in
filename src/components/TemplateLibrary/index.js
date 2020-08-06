import React, { useState, useEffect } from 'react';
import { Loader } from 'semantic-ui-react';

import { Library as TemplateLibraryRenderer } from '@accordproject/ui-components';
import { TemplateLibrary, Template, Clause } from '@accordproject/cicero-core';

import renderNodes from '../../utils/CiceroMarkToOOXML';
import attachVariableChangeListener from '../../utils/AttachVariableChangeListener';
import VariableVisitor from '../../utils/VariableVisitor';

const XML_HEADER = '<?xml version="1.0" encoding="utf-8" ?>';
const CUSTOM_XML_NAMESPACE = 'https://accordproject.org/';

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
      setTemplates(templateIndex);
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

  useEffect(() => {
    async function initializeDocument() {
      Office.context.document.customXmlParts.getByNamespaceAsync(CUSTOM_XML_NAMESPACE, result => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          if (result.value.length > 0) {
            const customXmlPart = result.value[0];
            customXmlPart.getNodesAsync('*/*', async result => {
              if (result.status === Office.AsyncResultStatus.Succeeded) {
                for (let index=0; index<result.value.length; ++index) {
                  const namespaceUri = result.value[index].namespaceUri;
                  const templateIndex = templates[namespaceUri];
                  const template = await Template.fromUrl(templateIndex.ciceroUrl);
                  const ciceroMark = templateToCiceroMark(template);
                  const numeration = VariableVisitor.getVariables(ciceroMark);
                  Word.run(async context => {
                    const contentControls = context.document.body.contentControls;
                    contentControls.load(['items/length', 'title']);
                    await context.sync();
                    for (let index=0; index<contentControls.items.length; ++index) {
                      if (numeration.includes(contentControls.items[index].title)) {
                        attachVariableChangeListener(contentControls.items[index].title);
                      }
                    }
                  });
                }
              }
            });
          }
        }
      });
    }
    if (templates !== null) {
      initializeDocument();
    }
  }, [templates]);

  const setup = async ciceroMark => {
    await Word.run(async context => {
      let counter = {...overallCounter};
      context.document.body.insertBreak(Word.BreakType.line, Word.InsertLocation.end);
      ciceroMark.nodes.forEach(node => {
        renderNodes(context, node, counter);
      });
      await context.sync();
      setOverallCounter({
        ...overallCounter,
        ...counter,
      });
      for (const variableText in counter) {
        for (let index=1; index<=counter[variableText]; ++index) {
          attachVariableChangeListener(`${variableText.toUpperCase()[0]}${variableText.substring(1)}${index}`);
        }
      }
    });
  };

  const templateToCiceroMark = template => {
    const sampleText = template.getMetadata().getSample();
    const clause = new Clause(template);
    clause.parse(sampleText);
    const ciceroMark = clause.draft({ format : 'ciceromark_parsed' });
    return ciceroMark;
  };

  const loadTemplateText = async templateIndex => {
    // URL to compiled archive
    const template = await Template.fromUrl(templateIndex.ciceroUrl);
    const ciceroMark = templateToCiceroMark(template);
    setup(ciceroMark);
    const templateIdentifier = `${templateIndex.name}@${templateIndex.version}`;
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
          const customXmlPart = result.value[0];
          customXmlPart.getNodesAsync('*/*', result => {
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
      items = {Object.values(templates)}
      onPrimaryButtonClick={loadTemplateText}
      onSecondaryButtonClick={goToTemplateDetail}
      onUploadItem={onUploadTemplate}
    />
  );
};

export default LibraryComponent;
