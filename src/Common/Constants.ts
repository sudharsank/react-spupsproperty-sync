export const boolFormatter = `{
    "elmType": "div",
    "style": {
      "box-sizing": "border-box",
      "padding": "0 2px"
    },
    "attributes": {
      "class": {
        "operator": ":",
        "operands": [
          {
            "operator": "==",
            "operands": [
              "@currentField",
              true
            ]
          },
          "sp-css-backgroundColor-successBackground",
          {
            "operator": ":",
            "operands": [
              {
                "operator": "==",
                "operands": [
                  "@currentField",
                  false
                ]
              },
              "sp-css-backgroundColor-errorBackground",
              ""
            ]
          }
        ]
      }
    },
    "children": [
      {
        "elmType": "span",
        "attributes": {
          "iconName": {
            "operator": ":",
            "operands": [
              {
                "operator": "==",
                "operands": [
                  "@currentField",
                  true
                ]
              },
              "",
              {
                "operator": ":",
                "operands": [
                  {
                    "operator": "==",
                    "operands": [
                      "@currentField",
                      false
                    ]
                  },
                  "",
                  ""
                ]
              }
            ]
          },
          "class": {
            "operator": ":",
            "operands": [
              {
                "operator": "==",
                "operands": [
                  "@currentField",
                  true
                ]
              },
              "",
              {
                "operator": ":",
                "operands": [
                  {
                    "operator": "==",
                    "operands": [
                      "@currentField",
                      false
                    ]
                  },
                  "",
                  ""
                ]
              }
            ]
          }
        }
      },
      {
        "elmType": "span",
        "style": {
          "padding": "0 2px"
        },
        "txtContent": "=if(@currentField==true,'Yes',if(@currentField==false,'No',''))",
        "attributes": {
          "class": {
            "operator": ":",
            "operands": [
              {
                "operator": "==",
                "operands": [
                  "@currentField",
                  true
                ]
              },
              "",
              {
                "operator": ":",
                "operands": [
                  {
                    "operator": "==",
                    "operands": [
                      "@currentField",
                      false
                    ]
                  },
                  "",
                  ""
                ]
              }
            ]
          }
        }
      }
    ]
  }`;