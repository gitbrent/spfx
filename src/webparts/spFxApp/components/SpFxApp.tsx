/**
* ORIG: https://github.com/SharePoint/sp-dev-fx-controls-react/tree/master/src/controls/listItemPicker
* ADDED: Ability to filter List query using `listFilter` prop
*/
import * as React from "react";
import styles from "./SpFxApp.module.scss";
import { ISpFxAppProps } from "./ISpFxAppProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { ListItemPicker } from '../components/listItemPicker/ListItemPicker';

export default class SpFxApp extends React.Component<ISpFxAppProps, {}> {
  public render(): React.ReactElement<ISpFxAppProps> {
    return (
      <div className={styles.spFxApp}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.title}>Welcome to SharePoint!</span>
              <p className={styles.subTitle}>
                Customize SharePoint experiences using Web Parts.
              </p>
              <p className={styles.description}>
                {escape(this.props.description)}
              </p>
              <a href="https://aka.ms/spfx" className={styles.button}>
                <span className={styles.label}>Learn more</span>
              </a>
            </div>
          </div>
        </div>

		<ListItemPicker
			listId='12a01bb8-1234-ABCD-b4e9-00dde9c8ef90'
			columnInternalName='Title'
			listFilter='Archived_x003f_ eq 1'
			itemLimit={1}
			onSelectedItem={this.onSelectedItem}
			context={this.props.context}
			suggestionsHeaderText='Select Account Name'
		/>

      </div>
    );
  }
}
