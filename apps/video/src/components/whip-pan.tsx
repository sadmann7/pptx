import type {
  TransitionPresentation,
  TransitionPresentationComponentProps,
} from "@remotion/transitions";
import { AbsoluteFill, interpolate, useVideoConfig } from "remotion";

type WhipPanProps = Record<string, unknown> & {
  direction?: "left" | "right";
  blur?: number;
};

function WhipPanPresenting({
  children,
  presentationDirection,
  presentationProgress,
  passedProps,
}: TransitionPresentationComponentProps<WhipPanProps>) {
  const { width } = useVideoConfig();
  const { direction = "left", blur = 20 } = passedProps;
  const sign = direction === "left" ? -1 : 1;

  const isExiting = presentationDirection === "exiting";
  const progress = presentationProgress;

  const translateX = isExiting
    ? interpolate(progress, [0, 1], [0, sign * width])
    : interpolate(progress, [0, 1], [-sign * width, 0]);

  const blurAmount = interpolate(progress, [0, 0.3, 0.7, 1], [0, blur, blur, 0], {
    extrapolateLeft: "clamp",
    extrapolateRight: "clamp",
  });

  const scaleX =
    1 +
    interpolate(progress, [0, 0.5, 1], [0, 0.02, 0], {
      extrapolateLeft: "clamp",
      extrapolateRight: "clamp",
    });

  return (
    <AbsoluteFill
      style={{
        translate: `${translateX}px 0px`,
        filter: `blur(${blurAmount}px)`,
        scale: `${scaleX} 1`,
      }}
    >
      {children}
    </AbsoluteFill>
  );
}

export function whipPan(props: WhipPanProps = {}): TransitionPresentation<WhipPanProps> {
  return {
    component: WhipPanPresenting,
    props,
  };
}
