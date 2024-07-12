import * as React from "react";
import styles from "./GroupMembers.module.scss";
import type { IGroupMembersProps } from "./IGroupMembersProps";
import { PrimaryButton } from "@fluentui/react/lib/Button";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { IMember } from "../models/IMember";
import { LivePersonaCard } from "./LivePersonaCard/LivePersonaCard";
import { Persona, PersonaSize } from "@fluentui/react/lib/Persona";

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
          <div className={styles.memberContainer}>
            {this.state.members.map((member) => {
              return (
                <LivePersonaCard
                  key={member.id}
                  upn={member.userPrincipalName}
                  serviceScope={this.props.context.serviceScope}
                  template={
                    <Persona
                      imageUrl={`/_layouts/15/userphoto.aspx?size=L&accountname=${member.mail}`}
                      imageShouldFadeIn={false}
                      imageShouldStartVisible={true}
                      text={member.displayName}
                      secondaryText={member.jobTitle}
                      tertiaryText={member.mail}
                      size={PersonaSize.size48}
                      imageAlt={member.displayName}
                    />
                  }
                />
              );
            })}
          </div>
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
