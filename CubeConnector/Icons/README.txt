CUBECONNECTOR CUSTOM ICONS
===========================

This folder contains custom icons for the CubeConnector ribbon.

FOLDER STRUCTURE:
-----------------
  Light/  - Icons for Light mode (dark colored icons on light backgrounds)
  Dark/   - Icons for Dark mode (light colored icons on dark backgrounds)

ICON REQUIREMENTS:
------------------
  - Format: PNG with transparent background
  - Size: 32x32 pixels (for large buttons)
  - Color scheme:
    * Light mode icons: Use dark colors (#333333 or similar)
    * Dark mode icons: Use light colors (#FFFFFF or similar)

REQUIRED ICON FILES:
--------------------
Place the following icons in BOTH Light/ and Dark/ folders:

  icon_connect.png         - Database/connection icon
  icon_refresh_all.png     - Full refresh/sync icon
  icon_refresh_sheet.png   - Sheet/page refresh icon
  icon_refresh_selection.png - Selection/cell refresh icon
  icon_auto_refresh.png    - Automatic/timer refresh icon
  icon_clear_cache.png     - Clear/delete cache icon
  icon_drill_detail.png    - Drill down/detail icon
  icon_drill_pivot.png     - Pivot table icon
  icon_help.png            - Help/question mark icon

FALLBACK BEHAVIOR:
------------------
If custom icons are not found, the ribbon will automatically fall back to
Office built-in icons (imageMso). This means the ribbon will work perfectly
even without custom icons - they're optional for branding purposes.

TESTING:
--------
To test your icons:
1. Place your PNG files in the Light/ and Dark/ folders
2. Rebuild the project
3. Load the add-in in Excel
4. Switch between Light and Dark Office themes to verify both icon sets

DESIGN TIPS:
------------
- Keep designs simple and recognizable at small sizes
- Use consistent line weights across all icons
- Maintain visual cohesion with the existing Excel ribbon style
- Test icons in both light and dark modes before finalizing
