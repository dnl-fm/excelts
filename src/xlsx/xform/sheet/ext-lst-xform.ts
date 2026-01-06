/* eslint-disable max-classes-per-file */

import CompositeXform from '../composite-xform.ts';
import ConditionalFormattingsExt from './cf-ext/conditional-formattings-ext-xform.ts';
import BaseXform from '../base-xform.ts';
import type { XmlStreamWriter } from '../xform-types.ts';

type CfRuleExtModel = {
  type?: string;
  x14Id?: string;
  priority?: number;
  custom?: boolean;
  dataBar?: unknown;
  iconSet?: string;
  gradient?: boolean;
  [key: string]: unknown;
};

type ConditionalFormattingExtModel = {
  ref?: string;
  rules: CfRuleExtModel[];
};

type ConditionalFormattingsModel = ConditionalFormattingExtModel[] & {
  hasExtContent?: boolean;
};

interface ExtModel {
  conditionalFormattings?: ConditionalFormattingsModel;
  [key: string]: unknown;
}

class ExtXform extends CompositeXform {
  conditionalFormattings: ConditionalFormattingsExt;

  constructor() {
    super();
    this.conditionalFormattings = new ConditionalFormattingsExt();
    this.map = {
      'x14:conditionalFormattings': this.conditionalFormattings,
    };
  }

  get tag(): string {
    return 'ext';
  }

  hasContent(model: ExtModel): boolean {
    return this.conditionalFormattings.hasContent(model.conditionalFormattings);
  }

  prepare(model: ExtModel, options: unknown): void {
    this.conditionalFormattings.prepare(model.conditionalFormattings, options);
  }

  render(xmlStream: XmlStreamWriter, model: ExtModel): void {
    xmlStream.openNode('ext', {
      uri: '{78C0D931-6437-407d-A8EE-F0AAD7539E65}',
      'xmlns:x14': 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main',
    });

    this.conditionalFormattings.render(xmlStream, model.conditionalFormattings);

    xmlStream.closeNode();
  }

  createNewModel(): ExtModel {
    return {};
  }

  onParserClose(name: string, parser: BaseXform): void {
    (this.model as ExtModel)[name] = parser.model;
  }
}

class ExtLstXform extends CompositeXform {
  ext: ExtXform;

  constructor() {
    super();
    this.ext = new ExtXform();
    this.map = {
      ext: this.ext,
    };
  }

  get tag(): string {
    return 'extLst';
  }

  prepare(model: ExtModel, options: unknown): void {
    this.ext.prepare(model, options);
  }

  hasContent(model: ExtModel): boolean {
    return this.ext.hasContent(model);
  }

  render(xmlStream: XmlStreamWriter, model: ExtModel): void {
    if (!this.hasContent(model)) {
      return;
    }

    xmlStream.openNode('extLst');
    this.ext.render(xmlStream, model);
    xmlStream.closeNode();
  }

  createNewModel(): ExtModel {
    return {};
  }

  onParserClose(_name: string, parser: BaseXform): void {
    Object.assign(this.model as ExtModel, parser.model);
  }
}

export default ExtLstXform;