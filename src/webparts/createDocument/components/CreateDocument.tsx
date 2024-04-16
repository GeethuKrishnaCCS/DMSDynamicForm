import * as React from 'react';
import styles from './CreateDocument.module.scss';
import { ICreateDocumentProps } from '../interfaces/ICreateDocumentProps';

export default class CreateDocument extends React.Component<ICreateDocumentProps, {}> {
  public render(): React.ReactElement<ICreateDocumentProps> {
   

    return (
      <section className={`${styles.createDocument}`}>
        
      </section>
    );
  }
}
