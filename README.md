# vba
More or less useful scripts in VBA (playground) 

- [Bookmarks](Bookmarks.vba) bunch of functions to handle bookmarks in Word
  - `GetBookmarkNames` returns collection with all bookmark names
  - `ReadTextFromBookmark`
  - `WriteTextToBookmark`
  - `ClearBookmark` clears bookmark content but left bookmark intact
  - `ClearAllBookmarks`
  - `RemoveBookmark` removes bookmark but left content intact
  - `RemoveAllBookmarks`
  - `RemoveBookmarkWithContent`
  - `RemoveAllBookmarkWithContent`
- [HexToVBAColor](HexToVBAColor.vba) converts normal hex color values to VBA type, `hexColor` parameter can be a string in any of following formats: `#ff0000`, `ff0000`, `#f00`, `f00`. (VBA constructs color codes in very odd way by joining the BGR codes, so `#aabbcc` becomes `&Hccbbaa`)
- [PoprawDaty](PoprawDaty.vba) normalize the date format to the ISO 8601:2004 standard
