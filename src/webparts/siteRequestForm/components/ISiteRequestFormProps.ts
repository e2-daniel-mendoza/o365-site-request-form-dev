import { WebPartContext } from "@microsoft/sp-webpart-base";
import {
  IReadonlyTheme
} from '@microsoft/sp-component-base';

export interface ISiteRequestFormProps {
  description: string;
  context: WebPartContext;
  optionsListTitle: string;
  flowURL: string;
  themeVariant: IReadonlyTheme;
}
