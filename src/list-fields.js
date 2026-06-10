'use strict';

function xmlEscape(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function splitChoices(raw) {
  return String(raw)
    .split(',')
    .map(choice => choice.trim())
    .filter(Boolean);
}

function buildCreateFieldSchemaXml(values) {
  const type = values.type;
  const name = values.name;
  const displayName = values['display-name'] || name;
  const attrs = [
    `Name="${xmlEscape(name)}"`,
    `DisplayName="${xmlEscape(displayName)}"`,
  ];
  if (values.required === true) {
    attrs.push('Required="TRUE"');
  }

  if (type === 'text') {
    attrs.push('Type="Text"');
    return `<Field ${attrs.join(' ')} />`;
  }

  if (type === 'datetime') {
    attrs.push('Type="DateTime"');
    const format = values.format === 'date-time' ? 'DateTime' : 'DateOnly';
    attrs.push(`Format="${format}"`);
    return `<Field ${attrs.join(' ')} />`;
  }

  if (type === 'choice') {
    if (values.choices === undefined || values.choices === '') {
      throw new Error('--choices is required when --type is choice');
    }
    const choices = splitChoices(values.choices);
    if (!choices.length) {
      throw new Error('--choices must contain at least one non-empty value');
    }
    attrs.push('Type="Choice"');
    const choiceXml = choices.map(choice => `<CHOICE>${xmlEscape(choice)}</CHOICE>`).join('');
    return `<Field ${attrs.join(' ')}><CHOICES>${choiceXml}</CHOICES></Field>`;
  }

  throw new Error(`Unsupported --type value: ${type}. Expected text, datetime, or choice.`);
}

function buildCreateFieldXmlBody(values) {
  const addToDefaultView = values['add-to-default-view'] !== false;
  return JSON.stringify({
    parameters: {
      __metadata: { type: 'SP.XmlSchemaFieldCreationInformation' },
      SchemaXml: buildCreateFieldSchemaXml(values),
      Options: addToDefaultView ? 8 : 0,
    },
  });
}

module.exports = {
  buildCreateFieldSchemaXml,
  buildCreateFieldXmlBody,
  splitChoices,
  xmlEscape,
};
