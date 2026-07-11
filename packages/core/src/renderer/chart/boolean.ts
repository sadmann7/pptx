import { parseOoxmlBool } from "../../ooxml/boolean";
import { SafeXmlNode } from "../../ooxml/xml";

export function parseOoxmlBoolElement(node: SafeXmlNode): boolean {
  if (!node.exists()) return false;
  return parseOoxmlBoolValue(node.attr("val"), true);
}

export function parseOoxmlBoolValue(value: string | undefined, defaultValue: boolean): boolean {
  return parseOoxmlBool(value, defaultValue);
}
