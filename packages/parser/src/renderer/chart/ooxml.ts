import { SafeXmlNode } from '../../parser/XmlParser';
import { parseOoxmlBool } from '../../parser/booleans';

export function parseOoxmlBoolElement(node: SafeXmlNode): boolean {
  if (!node.exists()) return false;
  return parseOoxmlBoolValue(node.attr('val'), true);
}

export function parseOoxmlBoolValue(value: string | undefined, defaultValue: boolean): boolean {
  return parseOoxmlBool(value, defaultValue);
}
