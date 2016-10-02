import * as React from 'react';
import { Persona, PersonaPresence, PersonaSize } from 'office-ui-fabric-react';
import {IContactsDataService} from '../../../shared/IContactsDataService';

import styles from '../ReactContacts.module.scss';
import { IContactsWebPartProps } from '../../contacts/IContactsWebPartProps';
import {IContact} from '../../../shared/IContact';

export interface IReactContactsProps extends IContactsWebPartProps {
  dataService: IContactsDataService;
}

export interface IReactContactsState {
  contacts: IContact[];
}

export default class ReactContacts extends React.Component<IReactContactsProps, IReactContactsState> {
  constructor(props: IReactContactsProps, state: IReactContactsState) {
    super(props);
    this.state = {
      contacts: []
    };
  }

  public render(): JSX.Element {
    const contacts: JSX.Element[] = this.state.contacts.
      map(c => {
        return <Persona
          primaryText={c.firstName + ' ' + c.lastName}
          imageUrl={c.photoUrl}
          title={c.jobTitle}
          presence={ PersonaPresence.online }
          size={this.props.large ? PersonaSize.large : PersonaSize.small }
          >
        </Persona>;
      });

    return (
      <div className={styles.reactContacts}>
        <div className={styles.container}>
          <h1 class='ms-fontSize-xxl'> {this.props.title}</h1>
          {contacts}
        </div>
      </div>
    );
  }

  private loadContacts(): void {
    this.state.contacts = [];
    this.props.dataService.loadData(this.props.numberOfContacts, this.props.listName).then(response => {
      this.setState({
        contacts: response.sort(this.props.sortAscending ? this.sortContact : this.sortContactReverse)
      });
    });
  }

  public componentDidMount(): void {
    this.loadContacts();
  }

  public componentDidUpdate(prevProps: IReactContactsProps, prevState: IReactContactsState, prevContext: any): void {
    if (prevProps.large != this.props.large ||
      prevProps.listName != this.props.listName ||
      prevProps.numberOfContacts != this.props.numberOfContacts ||
      prevProps.sortAscending != this.props.sortAscending ||
      prevProps.title != this.props.title) {
      this.loadContacts();
    }
  }

  protected sortContact(a: IContact, b: IContact): number {
    if (a.lastName > b.lastName) {
      return 1;
    } else if (a.lastName < b.lastName) {
      return -1;
    }
    return 0;
  }
  protected sortContactReverse(a: IContact, b: IContact): number {
    if (a.lastName > b.lastName) {
      return -1;
    } else if (a.lastName < b.lastName) {
      return 1;
    }
    return 0;
  }
}
