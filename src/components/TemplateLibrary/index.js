import React, { useState, useEffect } from 'react';
import { Loader } from 'semantic-ui-react';
import { toast } from 'react-semantic-toasts';

import { Library as TemplateLibraryRenderer } from '@accordproject/ui-components';
import { TemplateLibrary, Template, Clause } from '@accordproject/cicero-core';

import ooxmlGenerator from '../../utils/CiceroMarkToOOXML';
import attachVariableChangeListener from '../../utils/AttachVariableChangeListener';
import VariableVisitor from '../../utils/VariableVisitor';
import titleGenerator from '../../utils/TitleGenerator';
import spec from '../../constants/spec';
import triggerClauseParse from '../../utils/TriggerClauseParse';

const CUSTOM_XML_NAMESPACE = 'https://accordproject.org/';
const XML_HEADER = '<?xml version="1.0" encoding="utf-8" ?>';

/**
 * Template library Renderer.
 *
 * @returns {React.FC} Template library
 */
const LibraryComponent = () => {
  const [templates, setTemplates] = useState(null);
  const [overallCounter, setOverallCounter] = useState({});

  useEffect(() => {
    /**
     * Loads the template library from https://templates.accordproject.org/ and stores them in the state.
     */
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

  /**
   * Renders an uploaded template.
   *
   * @param {MouseEvent} event event containing the file object
   */
  const onUploadTemplate = async event => {
    const fileUploaded = event.target.files[0];
    try {
      const template = await Template.fromArchive(fileUploaded);
      const ciceroMark = templateToCiceroMark(template);
      setup(ciceroMark, template);
    }
    catch (error) {
      Office.context.ui.displayDialogAsync(`${window.location.origin}/bad-file.html`, { width: 30, height: 8 });
    }
  };

  useEffect(() => {
    /**
     * Initializes the document by fetching the templates whose identifier is stored in CustomXMLPart.
     */
    async function initializeDocument() {
      Office.context.document.customXmlParts.getByNamespaceAsync(CUSTOM_XML_NAMESPACE, result => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          if (result.value.length > 0) {
            const customXmlPart = result.value[0];
            customXmlPart.getNodesAsync('*/*', async result => {
              if (result.status === Office.AsyncResultStatus.Succeeded) {
                for (let index=0; index<result.value.length; ++index) {
                  const templateIdentifier = result.value[index].namespaceUri;
                  const templateIndex = templates[templateIdentifier];
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
                    triggerClauseParse(templateIdentifier, template);
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


  /**
   * Sets up a template for rendering.
   *
   * @param {object} ciceroMark Ciceromark JSON
   * @param {object} template   Template object
   */
  const setup = async (ciceroMark, template) => {
    await Word.run(async context => {
      let counter = { ...overallCounter };
      let ooxml = ooxmlGenerator(ciceroMark, counter, '');
      const templateIdentifier = template.getIdentifier();
      ooxml = `
        <w:sdt>
          <w:sdtPr>
            <w:lock w:val="contentLocked" />
            <w:alias w:val="${templateIdentifier}"/>
          </w:sdtPr>
          <w:sdtContent>
          ${ooxml}
          </w:sdtContent>
        </w:sdt>
      `;
      ooxml = `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
      <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
        <pkg:xmlData>
          <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
          </Relationships>
        </pkg:xmlData>
      </pkg:part>
      <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
        <pkg:xmlData>
          <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" >
          ${ooxml}
          <w:p />
          </w:document>
        </pkg:xmlData>
      </pkg:part>
      ${spec}
      </pkg:package>`;
      context.document.body.insertOoxml(ooxml, Word.InsertLocation.end);
      await context.sync();
      setOverallCounter({
        ...overallCounter,
        ...counter,
      });
      for (const variableText in counter) {
        for (let index=1; index<=counter[variableText].count; ++index) {
          attachVariableChangeListener(
            titleGenerator(`${variableText.toUpperCase()[0]}${variableText.substring(1)}${index}`, counter[variableText].type)
          );
        }
      }
      triggerClauseParse(templateIdentifier, template);
    });
  };

  /**
   * Converts a template text to CiceroMark JSON.
   *
   * @param {object} template The template object
   * @returns {object} CiceroMark JSON of a template
   */
  const templateToCiceroMark = template => {
    const sampleText = template.getMetadata().getSample();
    const clause = new Clause(template);
    clause.parse(sampleText);
    const ciceroMark = clause.draft({ format : 'ciceromark_parsed' });
    return ciceroMark;
  };

  /**
   * Fetches templateIndex from https://templates.accordproject.org/, load the template, and save template details to CustomXML.
   *
   * @param {object} templateIndex Details of a particular template like URL, author, displayName, etc.
   */
  const loadTemplateText = async templateIndex => {
    // URL to compiled archive
    const template = await Template.fromUrl(templateIndex.ciceroUrl);
    const ciceroMark = templateToCiceroMark(template);
    const templateIdentifier = template.getIdentifier();
    saveTemplateToXml(ciceroMark, template, templateIdentifier);
  };

  /**
   * Sets up the template for rendering and saves it to the XML if the identifier of template is unique.
   *
   * @param {object} ciceroMark         CiceroMark JSON of a template
   * @param {object} template           Template object
   * @param {string} templateIdentifier Identifier for a template
   */
  const saveTemplateToXml = (ciceroMark, template, templateIdentifier) => {
    Office.context.document.customXmlParts.getByNamespaceAsync(CUSTOM_XML_NAMESPACE, result => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        if (result.value.length === 0) {
          setup(ciceroMark, template);
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
              let identifierExists = false;
              let newXml = XML_HEADER + `<templates xmlns="${CUSTOM_XML_NAMESPACE}">`;
              if (result.value.length > 0) {
                for (let node=0; node < result.value.length; ++node) {
                  if (result.value[node].namespaceUri !== templateIdentifier) {
                    newXml += `<template xmlns="${result.value[node].namespaceUri}" />`;
                  }
                  else {
                    identifierExists = true;
                  }
                }
              }
              if(!identifierExists){
                setup(ciceroMark, template);
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
              }else{
                toast(
                  {
                    title: 'Duplicate template',
                    description: <p>Template cannot be inserted as it is already present in the document.</p>,
                    type: 'error',
                    time: 5000,
                    animation: 'fly down',
                  },
                );
              }
            }
          });
        }
      }
    });
  };

  /**
   * Redirects to the template URL.
   *
   * @param {object} template Template object
   */
  const goToTemplateDetail = template => {
    const templateOrigin = new URL(template.url).origin;
    const { name, version } = template;
    window.open(`${templateOrigin}/${name}@${version}.html`, '_blank');
  };

  if (!templates) {
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
