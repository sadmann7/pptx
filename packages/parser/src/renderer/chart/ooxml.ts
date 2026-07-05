import { parseOoxmlBool } from "../../ooxml/booleans";
import { SafeXmlNode } from "../../ooxml/xml-parser";

export function parseOoxmlBoolElement(node: SafeXmlNode): boolean {
  if (!node.exists()) return false;
  return parseOoxmlBoolValue(node.attr("val"), true);
}

export function parseOoxmlBoolValue(value: string | undefined, defaultValue: boolean): boolean {
  return parseOoxmlBool(value, defaultValue);
}
