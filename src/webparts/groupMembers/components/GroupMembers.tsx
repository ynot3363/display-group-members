import * as React from "react";
import styles from "./GroupMembers.module.scss";
import type { IGroupMembersProps } from "./IGroupMembersProps";
import { PrimaryButton } from "@fluentui/react/lib/Button";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { IMember } from "../models/IMember";

interface IGroupMembersState {
  members: IMember[];
}

export default class GroupMembers extends React.Component<
  IGroupMembersProps,
  IGroupMembersState
> {
  constructor(props: IGroupMembersProps, state: IGroupMembersState) {
    super(props);

    // Initialize the state of the component
    this.state = {
      members: [],
    };
  }

  public render(): React.ReactElement<IGroupMembersProps> {
    const { groups } = this.props;

    return (
      <section className={styles.container}>
        <div>
          <h4 className={styles.header}>
            <strong>Groups:</strong>
          </h4>
          <ul>
            {groups.map((group) => {
              return (
                <li className={styles.group} key={group.id}>
                  <span className={styles.groupName}>{group.displayName}</span>
                  <PrimaryButton
                    onClick={() => {
                      this.getMembers(group.id).catch((error) => {
                        console.error(error);
                      });
                    }}
                  >
                    Get Members
                  </PrimaryButton>
                </li>
              );
            })}
          </ul>
        </div>
        <div>
          <h4 className={styles.header}>
            <strong>Members:</strong>
          </h4>
          {this.state.members.map((member) => {
            return (
              <li key={member.id}>
                <span>{member.displayName}</span>
              </li>
            );
          })}
        </div>
      </section>
    );
  }

  private async getMembers(id: string): Promise<void> {
    await this.props.context.msGraphClientFactory
      .getClient("3")
      .then(async (client: MSGraphClientV3) => {
        await client
          .api(`groups/${id}/members`)
          .version("v1.0")
          .get((err, res) => {
            if (err) {
              console.error(err);
              return;
            }

            // Map the JSON response to the output array
            this.setState({ members: res.value });
          });
      });
  }
}
