// If this needs updating, the styles JSON can be retrieved from a template using:
//
// await Word.run(async (context) => {
//   const retrievedStyles = context.application.retrieveStylesFromBase64(externalDocument);
//   await context.sync();

//   console.log("Styles from the other document:", retrievedStyles.value);
// });

export const CustomStylesJson: string = `{
  "styles": [
    {
      "baseStyle": "Normal",
      "builtIn": false,
      "inUse": true,
      "linked": true,
      "nextParagraphStyle": "Codify",
      "nameLocal": "Codify",
      "priority": 1,
      "quickStyle": true,
      "type": "Paragraph",
      "unhideWhenUsed": false,
      "visibility": false,
      "paragraphFormat": {
        "alignment": "Left",
        "firstLineIndent": 0.0,
        "keepTogether": false,
        "keepWithNext": false,
        "leftIndent": 0.0,
        "lineSpacing": 12.0,
        "lineUnitAfter": 0.0,
        "lineUnitBefore": 0.0,
        "mirrorIndents": false,
        "outlineLevel": "OutlineLevelBodyText",
        "rightIndent": 0.0,
        "spaceAfter": 0.0,
        "spaceBefore": 0.0,
        "widowControl": true
      },
      "font": {
        "name": "Courier New",
        "size": 10.0,
        "bold": false,
        "italic": false,
        "color": "#000000",
        "underline": "None",
        "subscript": false,
        "superscript": false,
        "strikeThrough": false,
        "doubleStrikeThrough": false,
        "highlightColor": null,
        "hidden": false
      },
      "listTemplate": null,
      "tableStyle": null,
      "shading": {
        "backgroundPatternColor": "#EFF8FF",
        "foregroundPatternColor": "#000000",
        "texture": "None"
      },
      "borders": {
        "insideBorderColor": "#000000",
        "insideBorderType": "None",
        "insideBorderWidth": "None",
        "outsideBorderColor": "#000000",
        "outsideBorderType": "Single",
        "outsideBorderWidth": "Pt050",
        "items": [
          {
            "color": "#000000",
            "type": "Single",
            "width": "Pt050",
            "visible": true,
            "location": "Top"
          },
          {
            "color": "#000000",
            "type": "Single",
            "width": "Pt050",
            "visible": true,
            "location": "Left"
          },
          {
            "color": "#000000",
            "type": "Single",
            "width": "Pt050",
            "visible": true,
            "location": "Bottom"
          },
          {
            "color": "#000000",
            "type": "Single",
            "width": "Pt050",
            "visible": true,
            "location": "Right"
          },
          {
            "color": "#000000",
            "type": "None",
            "width": "None",
            "visible": false,
            "location": "InsideHorizontal"
          }
        ]
      }
    },
    {
      "baseStyle": "Table Normal",
      "builtIn": false,
      "inUse": true,
      "linked": false,
      "nextParagraphStyle": "MrBl4ck",
      "nameLocal": "MrBl4ck",
      "priority": 100,
      "quickStyle": false,
      "type": "Table",
      "unhideWhenUsed": false,
      "visibility": false,
      "paragraphFormat": {
        "alignment": "Left",
        "firstLineIndent": 0.0,
        "keepTogether": false,
        "keepWithNext": false,
        "leftIndent": 0.0,
        "lineSpacing": 12.0,
        "lineUnitAfter": 0.0,
        "lineUnitBefore": 0.0,
        "mirrorIndents": false,
        "outlineLevel": "OutlineLevelBodyText",
        "rightIndent": 0.0,
        "spaceAfter": 0.0,
        "spaceBefore": 0.0,
        "widowControl": true
      },
      "font": {
        "name": "Calibri",
        "size": 11.0,
        "bold": false,
        "italic": false,
        "color": "#000000",
        "underline": "None",
        "subscript": false,
        "superscript": false,
        "strikeThrough": false,
        "doubleStrikeThrough": false,
        "highlightColor": null,
        "hidden": false
      },
      "listTemplate": null,
      "tableStyle": {
        "allowBreakAcrossPage": false,
        "alignment": "Left",
        "bottomCellMargin": 0.0,
        "leftCellMargin": 0.08,
        "rightCellMargin": 0.08,
        "topCellMargin": 0.0,
        "cellSpacing": 0.0
      },
      "shading": {
        "backgroundPatternColor": null,
        "foregroundPatternColor": "#000000",
        "texture": "None"
      },
      "borders": {
        "insideBorderColor": "#000000",
        "insideBorderType": "Single",
        "insideBorderWidth": "Pt050",
        "outsideBorderColor": "#000000",
        "outsideBorderType": "Single",
        "outsideBorderWidth": "Pt050",
        "items": [
          {
            "color": "#000000",
            "type": "Single",
            "width": "Pt050",
            "visible": true,
            "location": "Top"
          },
          {
            "color": "#000000",
            "type": "Single",
            "width": "Pt050",
            "visible": true,
            "location": "Left"
          },
          {
            "color": "#000000",
            "type": "Single",
            "width": "Pt050",
            "visible": true,
            "location": "Bottom"
          },
          {
            "color": "#000000",
            "type": "Single",
            "width": "Pt050",
            "visible": true,
            "location": "Right"
          },
          {
            "color": "#000000",
            "type": "Single",
            "width": "Pt050",
            "visible": true,
            "location": "InsideHorizontal"
          },
          {
            "color": "#000000",
            "type": "Single",
            "width": "Pt050",
            "visible": true,
            "location": "InsideVertical"
          }
        ]
      }
    },
    {
      "baseStyle": "Heading 3",
      "builtIn": false,
      "inUse": true,
      "linked": true,
      "nextParagraphStyle": "Issue Heading",
      "nameLocal": "Issue Heading",
      "priority": 1,
      "quickStyle": true,
      "type": "Paragraph",
      "unhideWhenUsed": false,
      "visibility": false,
      "paragraphFormat": {
        "alignment": "Left",
        "firstLineIndent": 0.0,
        "keepTogether": false,
        "keepWithNext": true,
        "leftIndent": 0.0,
        "lineSpacing": 12.0,
        "lineUnitAfter": 0.0,
        "lineUnitBefore": 0.0,
        "mirrorIndents": false,
        "outlineLevel": "OutlineLevel3",
        "rightIndent": 0.0,
        "spaceAfter": 2.0,
        "spaceBefore": 22.0,
        "widowControl": true
      },
      "font": {
        "name": "Calibri",
        "size": 11.0,
        "bold": true,
        "italic": false,
        "color": "#000000",
        "underline": "None",
        "subscript": false,
        "superscript": false,
        "strikeThrough": false,
        "doubleStrikeThrough": false,
        "highlightColor": null,
        "hidden": false
      },
      "listTemplate": null,
      "tableStyle": null,
      "shading": {
        "backgroundPatternColor": null,
        "foregroundPatternColor": "#000000",
        "texture": "None"
      },
      "borders": {
        "insideBorderColor": "#000000",
        "insideBorderType": "None",
        "insideBorderWidth": "None",
        "outsideBorderColor": "#000000",
        "outsideBorderType": "None",
        "outsideBorderWidth": "None",
        "items": [
          {
            "color": "#000000",
            "type": "None",
            "width": "None",
            "visible": false,
            "location": "Top"
          },
          {
            "color": "#000000",
            "type": "None",
            "width": "None",
            "visible": false,
            "location": "Left"
          },
          {
            "color": "#000000",
            "type": "None",
            "width": "None",
            "visible": false,
            "location": "Bottom"
          },
          {
            "color": "#000000",
            "type": "None",
            "width": "None",
            "visible": false,
            "location": "Right"
          },
          {
            "color": "#000000",
            "type": "None",
            "width": "None",
            "visible": false,
            "location": "InsideHorizontal"
          }
        ]
      }
    },
    {
      "baseStyle": "Normal",
      "builtIn": false,
      "inUse": true,
      "linked": true,
      "nextParagraphStyle": "Issue SubHeading",
      "nameLocal": "Issue SubHeading",
      "priority": 1,
      "quickStyle": true,
      "type": "Paragraph",
      "unhideWhenUsed": false,
      "visibility": false,
      "paragraphFormat": {
        "alignment": "Left",
        "firstLineIndent": 0.0,
        "keepTogether": false,
        "keepWithNext": true,
        "leftIndent": 0.0,
        "lineSpacing": 12.0,
        "lineUnitAfter": 0.0,
        "lineUnitBefore": 0.0,
        "mirrorIndents": false,
        "outlineLevel": "OutlineLevelBodyText",
        "rightIndent": 0.0,
        "spaceAfter": 0.0,
        "spaceBefore": 11.0,
        "widowControl": true
      },
      "font": {
        "name": "Calibri",
        "size": 10.5,
        "bold": true,
        "italic": false,
        "color": "#000000",
        "underline": "None",
        "subscript": false,
        "superscript": false,
        "strikeThrough": false,
        "doubleStrikeThrough": false,
        "highlightColor": null,
        "hidden": false
      },
      "listTemplate": null,
      "tableStyle": null,
      "shading": {
        "backgroundPatternColor": null,
        "foregroundPatternColor": "#000000",
        "texture": "None"
      },
      "borders": {
        "insideBorderColor": "#000000",
        "insideBorderType": "None",
        "insideBorderWidth": "None",
        "outsideBorderColor": "#000000",
        "outsideBorderType": "None",
        "outsideBorderWidth": "None",
        "items": [
          {
            "color": "#000000",
            "type": "None",
            "width": "None",
            "visible": false,
            "location": "Top"
          },
          {
            "color": "#000000",
            "type": "None",
            "width": "None",
            "visible": false,
            "location": "Left"
          },
          {
            "color": "#000000",
            "type": "None",
            "width": "None",
            "visible": false,
            "location": "Bottom"
          },
          {
            "color": "#000000",
            "type": "None",
            "width": "None",
            "visible": false,
            "location": "Right"
          },
          {
            "color": "#000000",
            "type": "None",
            "width": "None",
            "visible": false,
            "location": "InsideHorizontal"
          }
        ]
      }
    }
  ]
}
`