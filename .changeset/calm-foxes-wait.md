---
"@diceui/pptx": minor
---

Add an `initialFocus` prop to `ThumbnailList`, and stop capturing focus on load by default. A deck can arrive without being asked for, such as one fetched on mount, and focusing a thumbnail then moved the page's tab position into the list. Pass `initialFocus` to focus the active thumbnail, a ref to focus something else, or a function to decide per load.
