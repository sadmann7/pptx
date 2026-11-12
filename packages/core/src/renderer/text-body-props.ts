import type { TextBody } from "../model/nodes/shape";
import { SafeXmlNode } from "../ooxml/xml";

// Layout/master bodyPr applies only when the slide shape omits the child entirely.
export function getEffectiveBodyPrChild(
  textBody: TextBody | undefined,
  childName: string,
): SafeXmlNode | undefined {
  const own = textBody?.bodyProperties?.child(childName);
  if (own?.exists()) return own;

  const inherited = textBody?.layoutBodyProperties?.child(childName);
  if (inherited?.exists()) return inherited;

  return undefined;
}
