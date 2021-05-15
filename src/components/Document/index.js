import React, { useEffect } from 'react';
import PropTypes from 'prop-types';

import './index.module.css';

/**
 * Renders the inserted templates into the document.
 *
 * @param {object} props Properties
 * @returns {React.ReactNode} JSX
 */
const DocumentComponent = props => {
  const { insertedTemplates, deletionTemplate, setDeletionTemplate, setActiveNav } = props;

  useEffect(()=>{
    if (deletionTemplate) {
      setActiveNav('library');
    }
  }, [deletionTemplate]);

  /**
   * Converts the data into JSX components.
   *
   * @param {Array} data Data to be rendered
   * @returns {Array} JSX array to be rendered
   */
  const renderTemplates = data => {
    return data.map(d=><div key={d.identifier} className='templateCard'>
      <h3 className='cardBody'>{d.name}</h3>
      <p className='identifier'>{d.identifier}</p>
      <div className='cardAction'>
        <span onClick={()=>{setDeletionTemplate(d.identifier);}}>
         Remove Template
        </span>
      </div>
    </div>);
  };

  return (
    <div className={'fullWidth'}>
      {renderTemplates(insertedTemplates)}
    </div>
  );
};

DocumentComponent.propTypes = {
  insertedTemplates: PropTypes.array,
  deletionTemplate: PropTypes.string,
  setDeletionTemplate: PropTypes.func,
  setActiveNav: PropTypes.func,
};


export default DocumentComponent;
