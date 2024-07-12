import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IGroup } from "../models/IGroup";

export interface IGroupMembersProps {
  groups: IGroup[];
  context: WebPartContext;
}
