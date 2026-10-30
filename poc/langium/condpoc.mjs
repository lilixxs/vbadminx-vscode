// poc/langium/src/condpoc.ts
import { writeFileSync } from "node:fs";
import { EmptyFileSystem } from "langium";
import { createDefaultCoreModule, createDefaultSharedCoreModule, inject, URI } from "langium";
import { parseHelper } from "langium/test";

// poc/langium/src/generated/ast.ts
import * as langium from "langium";
var Assignment = {
  $type: "Assignment",
  target: "target",
  value: "value"
};
var CallStatement = {
  $type: "CallStatement",
  args: "args",
  callee: "callee"
};
var Declaration = {
  $type: "Declaration",
  name: "name",
  type: "type"
};
var DoStatement = {
  $type: "DoStatement",
  body: "body",
  condition: "condition"
};
var Expression = {
  $type: "Expression",
  inner: "inner",
  left: "left",
  operand: "operand",
  operands: "operands",
  operator: "operator",
  operators: "operators"
};
var FieldDeclaration = {
  $type: "FieldDeclaration",
  name: "name",
  type: "type"
};
var ForStatement = {
  $type: "ForStatement",
  body: "body",
  from: "from",
  nextVar: "nextVar",
  to: "to",
  var: "var"
};
var FunctionCall = {
  $type: "FunctionCall",
  args: "args",
  callee: "callee",
  inner: "inner",
  left: "left",
  operand: "operand",
  operands: "operands",
  operator: "operator",
  operators: "operators"
};
var IfStatement = {
  $type: "IfStatement",
  condition: "condition",
  elifBlocks: "elifBlocks",
  elifConditions: "elifConditions",
  elseBlock: "elseBlock",
  thenBlock: "thenBlock"
};
var Literal = {
  $type: "Literal",
  inner: "inner",
  left: "left",
  operand: "operand",
  operands: "operands",
  operator: "operator",
  operators: "operators",
  value: "value"
};
var Module = {
  $type: "Module",
  members: "members"
};
var ModuleDeclaration = {
  $type: "ModuleDeclaration",
  name: "name",
  type: "type"
};
var ModuleMember = {
  $type: "ModuleMember"
};
var OptionStatement = {
  $type: "OptionStatement",
  name: "name",
  value: "value"
};
var Parameter = {
  $type: "Parameter",
  name: "name",
  type: "type"
};
var Procedure = {
  $type: "Procedure",
  endType: "endType",
  name: "name",
  params: "params",
  procType: "procType",
  returnType: "returnType",
  statements: "statements"
};
function isProcedure(item) {
  return reflection.isInstance(item, Procedure.$type);
}
var Reference = {
  $type: "Reference",
  inner: "inner",
  left: "left",
  operand: "operand",
  operands: "operands",
  operator: "operator",
  operators: "operators",
  target: "target"
};
var Statement = {
  $type: "Statement"
};
var TypeDeclaration = {
  $type: "TypeDeclaration",
  fields: "fields",
  name: "name"
};
var WhileStatement = {
  $type: "WhileStatement",
  body: "body",
  condition: "condition"
};
var WithStatement = {
  $type: "WithStatement",
  body: "body",
  target: "target"
};
var VbaMiniAstReflection = class extends langium.AbstractAstReflection {
  types = {
    Assignment: {
      name: Assignment.$type,
      properties: {
        target: {
          name: Assignment.target
        },
        value: {
          name: Assignment.value
        }
      },
      superTypes: [Statement.$type]
    },
    CallStatement: {
      name: CallStatement.$type,
      properties: {
        args: {
          name: CallStatement.args,
          defaultValue: [],
          optional: true
        },
        callee: {
          name: CallStatement.callee,
          referenceType: Procedure.$type
        }
      },
      superTypes: [Statement.$type]
    },
    Declaration: {
      name: Declaration.$type,
      properties: {
        name: {
          name: Declaration.name
        },
        type: {
          name: Declaration.type,
          optional: true
        }
      },
      superTypes: [Statement.$type]
    },
    DoStatement: {
      name: DoStatement.$type,
      properties: {
        body: {
          name: DoStatement.body,
          defaultValue: [],
          optional: true
        },
        condition: {
          name: DoStatement.condition,
          optional: true
        }
      },
      superTypes: [Statement.$type]
    },
    Expression: {
      name: Expression.$type,
      properties: {
        inner: {
          name: Expression.inner,
          optional: true
        },
        left: {
          name: Expression.left,
          optional: true
        },
        operand: {
          name: Expression.operand,
          optional: true
        },
        operands: {
          name: Expression.operands,
          defaultValue: [],
          optional: true
        },
        operator: {
          name: Expression.operator,
          optional: true
        },
        operators: {
          name: Expression.operators,
          defaultValue: [],
          optional: true
        }
      },
      superTypes: []
    },
    FieldDeclaration: {
      name: FieldDeclaration.$type,
      properties: {
        name: {
          name: FieldDeclaration.name
        },
        type: {
          name: FieldDeclaration.type,
          optional: true
        }
      },
      superTypes: []
    },
    ForStatement: {
      name: ForStatement.$type,
      properties: {
        body: {
          name: ForStatement.body,
          defaultValue: [],
          optional: true
        },
        from: {
          name: ForStatement.from
        },
        nextVar: {
          name: ForStatement.nextVar,
          optional: true
        },
        to: {
          name: ForStatement.to
        },
        var: {
          name: ForStatement.var
        }
      },
      superTypes: [Statement.$type]
    },
    FunctionCall: {
      name: FunctionCall.$type,
      properties: {
        args: {
          name: FunctionCall.args,
          defaultValue: [],
          optional: true
        },
        callee: {
          name: FunctionCall.callee,
          referenceType: Procedure.$type
        },
        inner: {
          name: FunctionCall.inner,
          optional: true
        },
        left: {
          name: FunctionCall.left,
          optional: true
        },
        operand: {
          name: FunctionCall.operand,
          optional: true
        },
        operands: {
          name: FunctionCall.operands,
          defaultValue: [],
          optional: true
        },
        operator: {
          name: FunctionCall.operator,
          optional: true
        },
        operators: {
          name: FunctionCall.operators,
          defaultValue: [],
          optional: true
        }
      },
      superTypes: [Expression.$type]
    },
    IfStatement: {
      name: IfStatement.$type,
      properties: {
        condition: {
          name: IfStatement.condition
        },
        elifBlocks: {
          name: IfStatement.elifBlocks,
          defaultValue: [],
          optional: true
        },
        elifConditions: {
          name: IfStatement.elifConditions,
          defaultValue: [],
          optional: true
        },
        elseBlock: {
          name: IfStatement.elseBlock,
          defaultValue: [],
          optional: true
        },
        thenBlock: {
          name: IfStatement.thenBlock,
          defaultValue: [],
          optional: true
        }
      },
      superTypes: [Statement.$type]
    },
    Literal: {
      name: Literal.$type,
      properties: {
        inner: {
          name: Literal.inner,
          optional: true
        },
        left: {
          name: Literal.left,
          optional: true
        },
        operand: {
          name: Literal.operand,
          optional: true
        },
        operands: {
          name: Literal.operands,
          defaultValue: [],
          optional: true
        },
        operator: {
          name: Literal.operator,
          optional: true
        },
        operators: {
          name: Literal.operators,
          defaultValue: [],
          optional: true
        },
        value: {
          name: Literal.value,
          defaultValue: false
        }
      },
      superTypes: [Expression.$type]
    },
    Module: {
      name: Module.$type,
      properties: {
        members: {
          name: Module.members,
          defaultValue: [],
          optional: true
        }
      },
      superTypes: []
    },
    ModuleDeclaration: {
      name: ModuleDeclaration.$type,
      properties: {
        name: {
          name: ModuleDeclaration.name
        },
        type: {
          name: ModuleDeclaration.type,
          optional: true
        }
      },
      superTypes: [ModuleMember.$type]
    },
    ModuleMember: {
      name: ModuleMember.$type,
      properties: {},
      superTypes: []
    },
    OptionStatement: {
      name: OptionStatement.$type,
      properties: {
        name: {
          name: OptionStatement.name
        },
        value: {
          name: OptionStatement.value,
          optional: true
        }
      },
      superTypes: [ModuleMember.$type]
    },
    Parameter: {
      name: Parameter.$type,
      properties: {
        name: {
          name: Parameter.name
        },
        type: {
          name: Parameter.type,
          optional: true
        }
      },
      superTypes: []
    },
    Procedure: {
      name: Procedure.$type,
      properties: {
        endType: {
          name: Procedure.endType
        },
        name: {
          name: Procedure.name
        },
        params: {
          name: Procedure.params,
          defaultValue: [],
          optional: true
        },
        procType: {
          name: Procedure.procType
        },
        returnType: {
          name: Procedure.returnType,
          optional: true
        },
        statements: {
          name: Procedure.statements,
          defaultValue: [],
          optional: true
        }
      },
      superTypes: [ModuleMember.$type]
    },
    Reference: {
      name: Reference.$type,
      properties: {
        inner: {
          name: Reference.inner,
          optional: true
        },
        left: {
          name: Reference.left,
          optional: true
        },
        operand: {
          name: Reference.operand,
          optional: true
        },
        operands: {
          name: Reference.operands,
          defaultValue: [],
          optional: true
        },
        operator: {
          name: Reference.operator,
          optional: true
        },
        operators: {
          name: Reference.operators,
          defaultValue: [],
          optional: true
        },
        target: {
          name: Reference.target
        }
      },
      superTypes: [Expression.$type]
    },
    Statement: {
      name: Statement.$type,
      properties: {},
      superTypes: []
    },
    TypeDeclaration: {
      name: TypeDeclaration.$type,
      properties: {
        fields: {
          name: TypeDeclaration.fields,
          defaultValue: []
        },
        name: {
          name: TypeDeclaration.name
        }
      },
      superTypes: [ModuleMember.$type]
    },
    WhileStatement: {
      name: WhileStatement.$type,
      properties: {
        body: {
          name: WhileStatement.body,
          defaultValue: [],
          optional: true
        },
        condition: {
          name: WhileStatement.condition
        }
      },
      superTypes: [Statement.$type]
    },
    WithStatement: {
      name: WithStatement.$type,
      properties: {
        body: {
          name: WithStatement.body,
          defaultValue: [],
          optional: true
        },
        target: {
          name: WithStatement.target
        }
      },
      superTypes: [Statement.$type]
    }
  };
};
var reflection = new VbaMiniAstReflection();

// poc/langium/src/generated/grammar.ts
import { loadGrammarFromJson } from "langium";
var loadedVbaMiniGrammar;
var VbaMiniGrammar = () => loadedVbaMiniGrammar ?? (loadedVbaMiniGrammar = loadGrammarFromJson(`{
  "$type": "Grammar",
  "isDeclared": true,
  "name": "VbaMini",
  "rules": [
    {
      "$type": "ParserRule",
      "entry": true,
      "name": "Module",
      "definition": {
        "$type": "Alternatives",
        "elements": [
          {
            "$type": "Assignment",
            "feature": "members",
            "operator": "+=",
            "terminal": {
              "$type": "RuleCall",
              "rule": {
                "$ref": "#/rules@1"
              },
              "arguments": []
            }
          },
          {
            "$type": "RuleCall",
            "rule": {
              "$ref": "#/rules@32"
            },
            "arguments": []
          }
        ],
        "cardinality": "*"
      },
      "fragment": false,
      "parameters": []
    },
    {
      "$type": "ParserRule",
      "name": "ModuleMember",
      "definition": {
        "$type": "Alternatives",
        "elements": [
          {
            "$type": "RuleCall",
            "rule": {
              "$ref": "#/rules@6"
            },
            "arguments": []
          },
          {
            "$type": "RuleCall",
            "rule": {
              "$ref": "#/rules@5"
            },
            "arguments": []
          },
          {
            "$type": "RuleCall",
            "rule": {
              "$ref": "#/rules@2"
            },
            "arguments": []
          },
          {
            "$type": "RuleCall",
            "rule": {
              "$ref": "#/rules@3"
            },
            "arguments": []
          }
        ]
      },
      "entry": false,
      "fragment": false,
      "parameters": []
    },
    {
      "$type": "ParserRule",
      "name": "OptionStatement",
      "definition": {
        "$type": "Group",
        "elements": [
          {
            "$type": "Keyword",
            "value": "Option"
          },
          {
            "$type": "Assignment",
            "feature": "name",
            "operator": "=",
            "terminal": {
              "$type": "RuleCall",
              "rule": {
                "$ref": "#/rules@27"
              },
              "arguments": []
            }
          },
          {
            "$type": "Assignment",
            "feature": "value",
            "operator": "=",
            "terminal": {
              "$type": "Alternatives",
              "elements": [
                {
                  "$type": "RuleCall",
                  "rule": {
                    "$ref": "#/rules@27"
                  },
                  "arguments": []
                },
                {
                  "$type": "RuleCall",
                  "rule": {
                    "$ref": "#/rules@29"
                  },
                  "arguments": []
                }
              ]
            },
            "cardinality": "?"
          }
        ]
      },
      "entry": false,
      "fragment": false,
      "parameters": []
    },
    {
      "$type": "ParserRule",
      "name": "TypeDeclaration",
      "definition": {
        "$type": "Group",
        "elements": [
          {
            "$type": "Alternatives",
            "elements": [
              {
                "$type": "Keyword",
                "value": "Public"
              },
              {
                "$type": "Keyword",
                "value": "Private"
              }
            ],
            "cardinality": "?"
          },
          {
            "$type": "Keyword",
            "value": "Type"
          },
          {
            "$type": "Assignment",
            "feature": "name",
            "operator": "=",
            "terminal": {
              "$type": "RuleCall",
              "rule": {
                "$ref": "#/rules@27"
              },
              "arguments": []
            }
          },
          {
            "$type": "RuleCall",
            "rule": {
              "$ref": "#/rules@32"
            },
            "arguments": []
          },
          {
            "$type": "Assignment",
            "feature": "fields",
            "operator": "+=",
            "terminal": {
              "$type": "RuleCall",
              "rule": {
                "$ref": "#/rules@4"
              },
              "arguments": []
            },
            "cardinality": "+"
          },
          {
            "$type": "Keyword",
            "value": "End"
          },
          {
            "$type": "Keyword",
            "value": "Type"
          }
        ]
      },
      "entry": false,
      "fragment": false,
      "parameters": []
    },
    {
      "$type": "ParserRule",
      "name": "FieldDeclaration",
      "definition": {
        "$type": "Group",
        "elements": [
          {
            "$type": "Assignment",
            "feature": "name",
            "operator": "=",
            "terminal": {
              "$type": "RuleCall",
              "rule": {
                "$ref": "#/rules@27"
              },
              "arguments": []
            }
          },
          {
            "$type": "Group",
            "elements": [
              {
                "$type": "Keyword",
                "value": "As"
              },
              {
                "$type": "Assignment",
                "feature": "type",
                "operator": "=",
                "terminal": {
                  "$type": "RuleCall",
                  "rule": {
                    "$ref": "#/rules@27"
                  },
                  "arguments": []
                }
              }
            ],
            "cardinality": "?"
          },
          {
            "$type": "RuleCall",
            "rule": {
              "$ref": "#/rules@32"
            },
            "arguments": []
          }
        ]
      },
      "entry": false,
      "fragment": false,
      "parameters": []
    },
    {
      "$type": "ParserRule",
      "name": "ModuleDeclaration",
      "definition": {
        "$type": "Group",
        "elements": [
          {
            "$type": "Alternatives",
            "elements": [
              {
                "$type": "Keyword",
                "value": "Public"
              },
              {
                "$type": "Keyword",
                "value": "Private"
              }
            ],
            "cardinality": "?"
          },
          {
            "$type": "Keyword",
            "value": "Dim"
          },
          {
            "$type": "Assignment",
            "feature": "name",
            "operator": "=",
            "terminal": {
              "$type": "RuleCall",
              "rule": {
                "$ref": "#/rules@27"
              },
              "arguments": []
            }
          },
          {
            "$type": "Group",
            "elements": [
              {
                "$type": "Keyword",
                "value": "As"
              },
              {
                "$type": "Assignment",
                "feature": "type",
                "operator": "=",
                "terminal": {
                  "$type": "RuleCall",
                  "rule": {
                    "$ref": "#/rules@27"
                  },
                  "arguments": []
                }
              }
            ],
            "cardinality": "?"
          }
        ]
      },
      "entry": false,
      "fragment": false,
      "parameters": []
    },
    {
      "$type": "ParserRule",
      "name": "Procedure",
      "definition": {
        "$type": "Group",
        "elements": [
          {
            "$type": "Alternatives",
            "elements": [
              {
                "$type": "Keyword",
                "value": "Public"
              },
              {
                "$type": "Keyword",
                "value": "Private"
              }
            ],
            "cardinality": "?"
          },
          {
            "$type": "Assignment",
            "feature": "procType",
            "operator": "=",
            "terminal": {
              "$type": "Alternatives",
              "elements": [
                {
                  "$type": "Keyword",
                  "value": "Sub"
                },
                {
                  "$type": "Keyword",
                  "value": "Function"
                }
              ]
            }
          },
          {
            "$type": "Assignment",
            "feature": "name",
            "operator": "=",
            "terminal": {
              "$type": "RuleCall",
              "rule": {
                "$ref": "#/rules@27"
              },
              "arguments": []
            }
          },
          {
            "$type": "Keyword",
            "value": "("
          },
          {
            "$type": "Group",
            "elements": [
              {
                "$type": "Assignment",
                "feature": "params",
                "operator": "+=",
                "terminal": {
                  "$type": "RuleCall",
                  "rule": {
                    "$ref": "#/rules@7"
                  },
                  "arguments": []
                }
              },
              {
                "$type": "Group",
                "elements": [
                  {
                    "$type": "Keyword",
                    "value": ","
                  },
                  {
                    "$type": "Assignment",
                    "feature": "params",
                    "operator": "+=",
                    "terminal": {
                      "$type": "RuleCall",
                      "rule": {
                        "$ref": "#/rules@7"
                      },
                      "arguments": []
                    }
                  }
                ],
                "cardinality": "*"
              }
            ],
            "cardinality": "?"
          },
          {
            "$type": "Keyword",
            "value": ")"
          },
          {
            "$type": "Group",
            "elements": [
              {
                "$type": "Keyword",
                "value": "As"
              },
              {
                "$type": "Assignment",
                "feature": "returnType",
                "operator": "=",
                "terminal": {
                  "$type": "RuleCall",
                  "rule": {
                    "$ref": "#/rules@27"
                  },
                  "arguments": []
                }
              }
            ],
            "cardinality": "?"
          },
          {
            "$type": "RuleCall",
            "rule": {
              "$ref": "#/rules@32"
            },
            "arguments": []
          },
          {
            "$type": "Assignment",
            "feature": "statements",
            "operator": "+=",
            "terminal": {
              "$type": "RuleCall",
              "rule": {
                "$ref": "#/rules@8"
              },
              "arguments": []
            },
            "cardinality": "*"
          },
          {
            "$type": "Keyword",
            "value": "End"
          },
          {
            "$type": "Assignment",
            "feature": "endType",
            "operator": "=",
            "terminal": {
              "$type": "Alternatives",
              "elements": [
                {
                  "$type": "Keyword",
                  "value": "Sub"
                },
                {
                  "$type": "Keyword",
                  "value": "Function"
                }
              ]
            }
          }
        ]
      },
      "entry": false,
      "fragment": false,
      "parameters": []
    },
    {
      "$type": "ParserRule",
      "name": "Parameter",
      "definition": {
        "$type": "Group",
        "elements": [
          {
            "$type": "Keyword",
            "value": "Optional",
            "cardinality": "?"
          },
          {
            "$type": "Alternatives",
            "elements": [
              {
                "$type": "Keyword",
                "value": "ByVal"
              },
              {
                "$type": "Keyword",
                "value": "ByRef"
              }
            ],
            "cardinality": "?"
          },
          {
            "$type": "Assignment",
            "feature": "name",
            "operator": "=",
            "terminal": {
              "$type": "RuleCall",
              "rule": {
                "$ref": "#/rules@27"
              },
              "arguments": []
            }
          },
          {
            "$type": "Group",
            "elements": [
              {
                "$type": "Keyword",
                "value": "As"
              },
              {
                "$type": "Assignment",
                "feature": "type",
                "operator": "=",
                "terminal": {
                  "$type": "RuleCall",
                  "rule": {
                    "$ref": "#/rules@27"
                  },
                  "arguments": []
                }
              }
            ],
            "cardinality": "?"
          }
        ]
      },
      "entry": false,
      "fragment": false,
      "parameters": []
    },
    {
      "$type": "ParserRule",
      "name": "Statement",
      "definition": {
        "$type": "Alternatives",
        "elements": [
          {
            "$type": "RuleCall",
            "rule": {
              "$ref": "#/rules@9"
            },
            "arguments": []
          },
          {
            "$type": "RuleCall",
            "rule": {
              "$ref": "#/rules@10"
            },
            "arguments": []
          },
          {
            "$type": "RuleCall",
            "rule": {
              "$ref": "#/rules@11"
            },
            "arguments": []
          },
          {
            "$type": "RuleCall",
            "rule": {
              "$ref": "#/rules@12"
            },
            "arguments": []
          },
          {
            "$type": "RuleCall",
            "rule": {
              "$ref": "#/rules@13"
            },
            "arguments": []
          },
          {
            "$type": "RuleCall",
            "rule": {
              "$ref": "#/rules@16"
            },
            "arguments": []
          },
          {
            "$type": "RuleCall",
            "rule": {
              "$ref": "#/rules@15"
            },
            "arguments": []
          },
          {
            "$type": "RuleCall",
            "rule": {
              "$ref": "#/rules@14"
            },
            "arguments": []
          },
          {
            "$type": "RuleCall",
            "rule": {
              "$ref": "#/rules@32"
            },
            "arguments": []
          }
        ]
      },
      "entry": false,
      "fragment": false,
      "parameters": []
    },
    {
      "$type": "ParserRule",
      "name": "Declaration",
      "definition": {
        "$type": "Group",
        "elements": [
          {
            "$type": "Alternatives",
            "elements": [
              {
                "$type": "Keyword",
                "value": "Dim"
              },
              {
                "$type": "Keyword",
                "value": "Static"
              }
            ]
          },
          {
            "$type": "Assignment",
            "feature": "name",
            "operator": "=",
            "terminal": {
              "$type": "RuleCall",
              "rule": {
                "$ref": "#/rules@27"
              },
              "arguments": []
            }
          },
          {
            "$type": "Group",
            "elements": [
              {
                "$type": "Keyword",
                "value": "As"
              },
              {
                "$type": "Assignment",
                "feature": "type",
                "operator": "=",
                "terminal": {
                  "$type": "RuleCall",
                  "rule": {
                    "$ref": "#/rules@27"
                  },
                  "arguments": []
                }
              }
            ],
            "cardinality": "?"
          }
        ]
      },
      "entry": false,
      "fragment": false,
      "parameters": []
    },
    {
      "$type": "ParserRule",
      "name": "Assignment",
      "definition": {
        "$type": "Group",
        "elements": [
          {
            "$type": "Alternatives",
            "elements": [
              {
                "$type": "Keyword",
                "value": "Let"
              },
              {
                "$type": "Keyword",
                "value": "Set"
              }
            ],
            "cardinality": "?"
          },
          {
            "$type": "Assignment",
            "feature": "target",
            "operator": "=",
            "terminal": {
              "$type": "RuleCall",
              "rule": {
                "$ref": "#/rules@26"
              },
              "arguments": []
            }
          },
          {
            "$type": "Keyword",
            "value": "="
          },
          {
            "$type": "Assignment",
            "feature": "value",
            "operator": "=",
            "terminal": {
              "$type": "RuleCall",
              "rule": {
                "$ref": "#/rules@17"
              },
              "arguments": []
            }
          }
        ]
      },
      "entry": false,
      "fragment": false,
      "parameters": []
    },
    {
      "$type": "ParserRule",
      "name": "CallStatement",
      "definition": {
        "$type": "Group",
        "elements": [
          {
            "$type": "Keyword",
            "value": "Call",
            "cardinality": "?"
          },
          {
            "$type": "Assignment",
            "feature": "callee",
            "operator": "=",
            "terminal": {
              "$type": "CrossReference",
              "type": {
                "$ref": "#/rules@6"
              },
              "terminal": {
                "$type": "RuleCall",
                "rule": {
                  "$ref": "#/rules@26"
                },
                "arguments": []
              },
              "deprecatedSyntax": false,
              "isMulti": false
            }
          },
          {
            "$type": "Keyword",
            "value": "("
          },
          {
            "$type": "Group",
            "elements": [
              {
                "$type": "Assignment",
                "feature": "args",
                "operator": "+=",
                "terminal": {
                  "$type": "RuleCall",
                  "rule": {
                    "$ref": "#/rules@17"
                  },
                  "arguments": []
                }
              },
              {
                "$type": "Group",
                "elements": [
                  {
                    "$type": "Keyword",
                    "value": ","
                  },
                  {
                    "$type": "Assignment",
                    "feature": "args",
                    "operator": "+=",
                    "terminal": {
                      "$type": "RuleCall",
                      "rule": {
                        "$ref": "#/rules@17"
                      },
                      "arguments": []
                    }
                  }
                ],
                "cardinality": "*"
              }
            ],
            "cardinality": "?"
          },
          {
            "$type": "Keyword",
            "value": ")"
          }
        ]
      },
      "entry": false,
      "fragment": false,
      "parameters": []
    },
    {
      "$type": "ParserRule",
      "name": "IfStatement",
      "definition": {
        "$type": "Group",
        "elements": [
          {
            "$type": "Keyword",
            "value": "If"
          },
          {
            "$type": "Assignment",
            "feature": "condition",
            "operator": "=",
            "terminal": {
              "$type": "RuleCall",
              "rule": {
                "$ref": "#/rules@17"
              },
              "arguments": []
            }
          },
          {
            "$type": "Keyword",
            "value": "Then"
          },
          {
            "$type": "RuleCall",
            "rule": {
              "$ref": "#/rules@32"
            },
            "arguments": []
          },
          {
            "$type": "Assignment",
            "feature": "thenBlock",
            "operator": "+=",
            "terminal": {
              "$type": "RuleCall",
              "rule": {
                "$ref": "#/rules@8"
              },
              "arguments": []
            },
            "cardinality": "*"
          },
          {
            "$type": "Group",
            "elements": [
              {
                "$type": "Keyword",
                "value": "ElseIf"
              },
              {
                "$type": "Assignment",
                "feature": "elifConditions",
                "operator": "+=",
                "terminal": {
                  "$type": "RuleCall",
                  "rule": {
                    "$ref": "#/rules@17"
                  },
                  "arguments": []
                }
              },
              {
                "$type": "Keyword",
                "value": "Then"
              },
              {
                "$type": "RuleCall",
                "rule": {
                  "$ref": "#/rules@32"
                },
                "arguments": []
              },
              {
                "$type": "Assignment",
                "feature": "elifBlocks",
                "operator": "+=",
                "terminal": {
                  "$type": "RuleCall",
                  "rule": {
                    "$ref": "#/rules@8"
                  },
                  "arguments": []
                },
                "cardinality": "*"
              }
            ],
            "cardinality": "*"
          },
          {
            "$type": "Group",
            "elements": [
              {
                "$type": "Keyword",
                "value": "Else"
              },
              {
                "$type": "RuleCall",
                "rule": {
                  "$ref": "#/rules@32"
                },
                "arguments": []
              },
              {
                "$type": "Assignment",
                "feature": "elseBlock",
                "operator": "+=",
                "terminal": {
                  "$type": "RuleCall",
                  "rule": {
                    "$ref": "#/rules@8"
                  },
                  "arguments": []
                },
                "cardinality": "*"
              }
            ],
            "cardinality": "?"
          },
          {
            "$type": "Keyword",
            "value": "End"
          },
          {
            "$type": "Keyword",
            "value": "If"
          }
        ]
      },
      "entry": false,
      "fragment": false,
      "parameters": []
    },
    {
      "$type": "ParserRule",
      "name": "ForStatement",
      "definition": {
        "$type": "Group",
        "elements": [
          {
            "$type": "Keyword",
            "value": "For"
          },
          {
            "$type": "Assignment",
            "feature": "var",
            "operator": "=",
            "terminal": {
              "$type": "RuleCall",
              "rule": {
                "$ref": "#/rules@27"
              },
              "arguments": []
            }
          },
          {
            "$type": "Keyword",
            "value": "="
          },
          {
            "$type": "Assignment",
            "feature": "from",
            "operator": "=",
            "terminal": {
              "$type": "RuleCall",
              "rule": {
                "$ref": "#/rules@17"
              },
              "arguments": []
            }
          },
          {
            "$type": "Keyword",
            "value": "To"
          },
          {
            "$type": "Assignment",
            "feature": "to",
            "operator": "=",
            "terminal": {
              "$type": "RuleCall",
              "rule": {
                "$ref": "#/rules@17"
              },
              "arguments": []
            }
          },
          {
            "$type": "RuleCall",
            "rule": {
              "$ref": "#/rules@32"
            },
            "arguments": []
          },
          {
            "$type": "Assignment",
            "feature": "body",
            "operator": "+=",
            "terminal": {
              "$type": "RuleCall",
              "rule": {
                "$ref": "#/rules@8"
              },
              "arguments": []
            },
            "cardinality": "*"
          },
          {
            "$type": "Keyword",
            "value": "Next"
          },
          {
            "$type": "Assignment",
            "feature": "nextVar",
            "operator": "=",
            "terminal": {
              "$type": "RuleCall",
              "rule": {
                "$ref": "#/rules@27"
              },
              "arguments": []
            },
            "cardinality": "?"
          }
        ]
      },
      "entry": false,
      "fragment": false,
      "parameters": []
    },
    {
      "$type": "ParserRule",
      "name": "WhileStatement",
      "definition": {
        "$type": "Group",
        "elements": [
          {
            "$type": "Keyword",
            "value": "While"
          },
          {
            "$type": "Assignment",
            "feature": "condition",
            "operator": "=",
            "terminal": {
              "$type": "RuleCall",
              "rule": {
                "$ref": "#/rules@17"
              },
              "arguments": []
            }
          },
          {
            "$type": "RuleCall",
            "rule": {
              "$ref": "#/rules@32"
            },
            "arguments": []
          },
          {
            "$type": "Assignment",
            "feature": "body",
            "operator": "+=",
            "terminal": {
              "$type": "RuleCall",
              "rule": {
                "$ref": "#/rules@8"
              },
              "arguments": []
            },
            "cardinality": "*"
          },
          {
            "$type": "Keyword",
            "value": "Wend"
          }
        ]
      },
      "entry": false,
      "fragment": false,
      "parameters": []
    },
    {
      "$type": "ParserRule",
      "name": "DoStatement",
      "definition": {
        "$type": "Group",
        "elements": [
          {
            "$type": "Keyword",
            "value": "Do"
          },
          {
            "$type": "Alternatives",
            "elements": [
              {
                "$type": "Keyword",
                "value": "While"
              },
              {
                "$type": "Keyword",
                "value": "Until"
              }
            ],
            "cardinality": "?"
          },
          {
            "$type": "Assignment",
            "feature": "condition",
            "operator": "=",
            "terminal": {
              "$type": "RuleCall",
              "rule": {
                "$ref": "#/rules@17"
              },
              "arguments": []
            },
            "cardinality": "?"
          },
          {
            "$type": "RuleCall",
            "rule": {
              "$ref": "#/rules@32"
            },
            "arguments": []
          },
          {
            "$type": "Assignment",
            "feature": "body",
            "operator": "+=",
            "terminal": {
              "$type": "RuleCall",
              "rule": {
                "$ref": "#/rules@8"
              },
              "arguments": []
            },
            "cardinality": "*"
          },
          {
            "$type": "Keyword",
            "value": "Loop"
          }
        ]
      },
      "entry": false,
      "fragment": false,
      "parameters": []
    },
    {
      "$type": "ParserRule",
      "name": "WithStatement",
      "definition": {
        "$type": "Group",
        "elements": [
          {
            "$type": "Keyword",
            "value": "With"
          },
          {
            "$type": "Assignment",
            "feature": "target",
            "operator": "=",
            "terminal": {
              "$type": "RuleCall",
              "rule": {
                "$ref": "#/rules@26"
              },
              "arguments": []
            }
          },
          {
            "$type": "RuleCall",
            "rule": {
              "$ref": "#/rules@32"
            },
            "arguments": []
          },
          {
            "$type": "Assignment",
            "feature": "body",
            "operator": "+=",
            "terminal": {
              "$type": "RuleCall",
              "rule": {
                "$ref": "#/rules@8"
              },
              "arguments": []
            },
            "cardinality": "*"
          },
          {
            "$type": "Keyword",
            "value": "End"
          },
          {
            "$type": "Keyword",
            "value": "With"
          }
        ]
      },
      "entry": false,
      "fragment": false,
      "parameters": []
    },
    {
      "$type": "ParserRule",
      "name": "Expression",
      "definition": {
        "$type": "Alternatives",
        "elements": [
          {
            "$type": "RuleCall",
            "rule": {
              "$ref": "#/rules@24"
            },
            "arguments": []
          },
          {
            "$type": "RuleCall",
            "rule": {
              "$ref": "#/rules@25"
            },
            "arguments": []
          },
          {
            "$type": "RuleCall",
            "rule": {
              "$ref": "#/rules@18"
            },
            "arguments": []
          },
          {
            "$type": "RuleCall",
            "rule": {
              "$ref": "#/rules@20"
            },
            "arguments": []
          },
          {
            "$type": "RuleCall",
            "rule": {
              "$ref": "#/rules@21"
            },
            "arguments": []
          }
        ]
      },
      "entry": false,
      "fragment": false,
      "parameters": []
    },
    {
      "$type": "ParserRule",
      "name": "BinaryExpression",
      "inferredType": {
        "$type": "InferredType",
        "name": "Expression"
      },
      "definition": {
        "$type": "Group",
        "elements": [
          {
            "$type": "Assignment",
            "feature": "left",
            "operator": "=",
            "terminal": {
              "$type": "RuleCall",
              "rule": {
                "$ref": "#/rules@22"
              },
              "arguments": []
            }
          },
          {
            "$type": "Group",
            "elements": [
              {
                "$type": "Assignment",
                "feature": "operators",
                "operator": "+=",
                "terminal": {
                  "$type": "RuleCall",
                  "rule": {
                    "$ref": "#/rules@19"
                  },
                  "arguments": []
                }
              },
              {
                "$type": "Assignment",
                "feature": "operands",
                "operator": "+=",
                "terminal": {
                  "$type": "RuleCall",
                  "rule": {
                    "$ref": "#/rules@22"
                  },
                  "arguments": []
                }
              }
            ],
            "cardinality": "*"
          }
        ]
      },
      "entry": false,
      "fragment": false,
      "parameters": []
    },
    {
      "$type": "ParserRule",
      "name": "BinaryOperator",
      "dataType": "string",
      "definition": {
        "$type": "Alternatives",
        "elements": [
          {
            "$type": "Keyword",
            "value": "+"
          },
          {
            "$type": "Keyword",
            "value": "-"
          },
          {
            "$type": "Keyword",
            "value": "*"
          },
          {
            "$type": "Keyword",
            "value": "/"
          },
          {
            "$type": "Keyword",
            "value": "&"
          },
          {
            "$type": "Keyword",
            "value": "="
          },
          {
            "$type": "Keyword",
            "value": "<"
          },
          {
            "$type": "Keyword",
            "value": ">"
          },
          {
            "$type": "Keyword",
            "value": "<="
          },
          {
            "$type": "Keyword",
            "value": ">="
          },
          {
            "$type": "Keyword",
            "value": "<>"
          },
          {
            "$type": "Keyword",
            "value": "And"
          },
          {
            "$type": "Keyword",
            "value": "Or"
          }
        ]
      },
      "entry": false,
      "fragment": false,
      "parameters": []
    },
    {
      "$type": "ParserRule",
      "name": "UnaryExpression",
      "inferredType": {
        "$type": "InferredType",
        "name": "Expression"
      },
      "definition": {
        "$type": "Group",
        "elements": [
          {
            "$type": "Assignment",
            "feature": "operator",
            "operator": "=",
            "terminal": {
              "$type": "Alternatives",
              "elements": [
                {
                  "$type": "Keyword",
                  "value": "-"
                },
                {
                  "$type": "Keyword",
                  "value": "Not"
                }
              ]
            }
          },
          {
            "$type": "Assignment",
            "feature": "operand",
            "operator": "=",
            "terminal": {
              "$type": "RuleCall",
              "rule": {
                "$ref": "#/rules@22"
              },
              "arguments": []
            }
          }
        ]
      },
      "entry": false,
      "fragment": false,
      "parameters": []
    },
    {
      "$type": "ParserRule",
      "name": "ParenthesizedExpression",
      "inferredType": {
        "$type": "InferredType",
        "name": "Expression"
      },
      "definition": {
        "$type": "Group",
        "elements": [
          {
            "$type": "Keyword",
            "value": "("
          },
          {
            "$type": "Assignment",
            "feature": "inner",
            "operator": "=",
            "terminal": {
              "$type": "RuleCall",
              "rule": {
                "$ref": "#/rules@17"
              },
              "arguments": []
            }
          },
          {
            "$type": "Keyword",
            "value": ")"
          }
        ]
      },
      "entry": false,
      "fragment": false,
      "parameters": []
    },
    {
      "$type": "ParserRule",
      "name": "PrimaryExpression",
      "inferredType": {
        "$type": "InferredType",
        "name": "Expression"
      },
      "definition": {
        "$type": "Alternatives",
        "elements": [
          {
            "$type": "RuleCall",
            "rule": {
              "$ref": "#/rules@24"
            },
            "arguments": []
          },
          {
            "$type": "RuleCall",
            "rule": {
              "$ref": "#/rules@23"
            },
            "arguments": []
          },
          {
            "$type": "RuleCall",
            "rule": {
              "$ref": "#/rules@25"
            },
            "arguments": []
          },
          {
            "$type": "RuleCall",
            "rule": {
              "$ref": "#/rules@21"
            },
            "arguments": []
          }
        ]
      },
      "entry": false,
      "fragment": false,
      "parameters": []
    },
    {
      "$type": "ParserRule",
      "name": "FunctionCall",
      "definition": {
        "$type": "Group",
        "elements": [
          {
            "$type": "Assignment",
            "feature": "callee",
            "operator": "=",
            "terminal": {
              "$type": "CrossReference",
              "type": {
                "$ref": "#/rules@6"
              },
              "terminal": {
                "$type": "RuleCall",
                "rule": {
                  "$ref": "#/rules@26"
                },
                "arguments": []
              },
              "deprecatedSyntax": false,
              "isMulti": false
            }
          },
          {
            "$type": "Keyword",
            "value": "("
          },
          {
            "$type": "Group",
            "elements": [
              {
                "$type": "Assignment",
                "feature": "args",
                "operator": "+=",
                "terminal": {
                  "$type": "RuleCall",
                  "rule": {
                    "$ref": "#/rules@17"
                  },
                  "arguments": []
                }
              },
              {
                "$type": "Group",
                "elements": [
                  {
                    "$type": "Keyword",
                    "value": ","
                  },
                  {
                    "$type": "Assignment",
                    "feature": "args",
                    "operator": "+=",
                    "terminal": {
                      "$type": "RuleCall",
                      "rule": {
                        "$ref": "#/rules@17"
                      },
                      "arguments": []
                    }
                  }
                ],
                "cardinality": "*"
              }
            ],
            "cardinality": "?"
          },
          {
            "$type": "Keyword",
            "value": ")"
          }
        ]
      },
      "entry": false,
      "fragment": false,
      "parameters": []
    },
    {
      "$type": "ParserRule",
      "name": "Literal",
      "definition": {
        "$type": "Assignment",
        "feature": "value",
        "operator": "=",
        "terminal": {
          "$type": "Alternatives",
          "elements": [
            {
              "$type": "RuleCall",
              "rule": {
                "$ref": "#/rules@28"
              },
              "arguments": []
            },
            {
              "$type": "RuleCall",
              "rule": {
                "$ref": "#/rules@29"
              },
              "arguments": []
            },
            {
              "$type": "RuleCall",
              "rule": {
                "$ref": "#/rules@30"
              },
              "arguments": []
            }
          ]
        }
      },
      "entry": false,
      "fragment": false,
      "parameters": []
    },
    {
      "$type": "ParserRule",
      "name": "Reference",
      "definition": {
        "$type": "Assignment",
        "feature": "target",
        "operator": "=",
        "terminal": {
          "$type": "RuleCall",
          "rule": {
            "$ref": "#/rules@26"
          },
          "arguments": []
        }
      },
      "entry": false,
      "fragment": false,
      "parameters": []
    },
    {
      "$type": "ParserRule",
      "name": "QualifiedName",
      "dataType": "string",
      "definition": {
        "$type": "Group",
        "elements": [
          {
            "$type": "RuleCall",
            "rule": {
              "$ref": "#/rules@27"
            },
            "arguments": []
          },
          {
            "$type": "Group",
            "elements": [
              {
                "$type": "Keyword",
                "value": "."
              },
              {
                "$type": "RuleCall",
                "rule": {
                  "$ref": "#/rules@27"
                },
                "arguments": []
              }
            ],
            "cardinality": "*"
          }
        ]
      },
      "entry": false,
      "fragment": false,
      "parameters": []
    },
    {
      "$type": "TerminalRule",
      "name": "ID",
      "definition": {
        "$type": "RegexToken",
        "regex": "/[_a-zA-Z][\\\\w_]*[\\\\$\\\\%\\\\&\\\\@\\\\^]?/",
        "parenthesized": false
      },
      "fragment": false,
      "hidden": false
    },
    {
      "$type": "TerminalRule",
      "name": "STRING",
      "definition": {
        "$type": "RegexToken",
        "regex": "/\\"([^\\"]|\\"\\")*\\"/",
        "parenthesized": false
      },
      "fragment": false,
      "hidden": false
    },
    {
      "$type": "TerminalRule",
      "name": "NUMBER",
      "type": {
        "$type": "ReturnType",
        "name": "number"
      },
      "definition": {
        "$type": "RegexToken",
        "regex": "/[0-9]+(\\\\.[0-9]+)?/",
        "parenthesized": false
      },
      "fragment": false,
      "hidden": false
    },
    {
      "$type": "TerminalRule",
      "name": "BOOLEAN",
      "type": {
        "$type": "ReturnType",
        "name": "boolean"
      },
      "definition": {
        "$type": "RegexToken",
        "regex": "/true|false/",
        "parenthesized": false
      },
      "fragment": false,
      "hidden": false
    },
    {
      "$type": "TerminalRule",
      "hidden": true,
      "name": "WS",
      "definition": {
        "$type": "RegexToken",
        "regex": "/[ \\\\t]+/",
        "parenthesized": false
      },
      "fragment": false
    },
    {
      "$type": "TerminalRule",
      "name": "NEWLINE",
      "definition": {
        "$type": "RegexToken",
        "regex": "/[\\\\r\\\\n]+/",
        "parenthesized": false
      },
      "fragment": false,
      "hidden": false
    },
    {
      "$type": "TerminalRule",
      "hidden": true,
      "name": "SL_COMMENT",
      "definition": {
        "$type": "RegexToken",
        "regex": "/'[^\\\\r\\\\n]*/",
        "parenthesized": false
      },
      "fragment": false
    }
  ],
  "imports": [],
  "interfaces": [],
  "types": []
}`));

// poc/langium/src/generated/module.ts
var VbaMiniLanguageMetaData = {
  languageId: "vbaMini",
  fileExtensions: [".bas"],
  caseInsensitive: true,
  mode: "development"
};
var VbaMiniGeneratedSharedModule = {
  AstReflection: () => new VbaMiniAstReflection()
};
var VbaMiniGeneratedModule = {
  Grammar: () => VbaMiniGrammar(),
  LanguageMetaData: () => VbaMiniLanguageMetaData,
  parser: {}
};

// poc/langium/src/condpoc.ts
function preprocess(text, environment) {
  const constants = { ...environment };
  const lines = text.split(/\r\n|\n|\r/);
  const stack = [];
  let activeLineCount = 0;
  let inactiveLineCount = 0;
  let unknownConstant = false;
  const evaluate = (expression) => {
    const match = /^(\w+)\s*(=|<>)\s*(\S+)$/.exec(expression.trim());
    if (match) {
      const [, name, operator, rawValue] = match;
      const value = constants[name];
      if (value === void 0) {
        unknownConstant = true;
        return void 0;
      }
      const literal = /^-?\d+$/.test(rawValue) ? Number(rawValue) : rawValue.replace(/^"|"$/g, "");
      return operator === "=" ? value === literal : value !== literal;
    }
    const sole = /^(\w+)$/.exec(expression.trim());
    if (sole) {
      const value = constants[sole[1]];
      if (value === void 0) {
        unknownConstant = true;
        return void 0;
      }
      return Boolean(value);
    }
    unknownConstant = true;
    return void 0;
  };
  const output = lines.map((line) => {
    const trimmed = line.trim();
    const constMatch = /^#Const\s+(\w+)\s*=\s*(.+)$/i.exec(trimmed);
    if (constMatch) {
      const raw = constMatch[2].trim();
      constants[constMatch[1]] = /^-?\d+$/.test(raw) ? Number(raw) : raw.replace(/^"|"$/g, "");
      return "";
    }
    const ifMatch = /^#If\s+(.+?)\s+Then$/i.exec(trimmed) ?? /^#If\s+(.+)$/i.exec(trimmed);
    if (ifMatch) {
      const result = evaluate(ifMatch[1]);
      const active2 = result === true;
      stack.push({ active: active2, taken: active2 });
      return "";
    }
    const elifMatch = /^#ElseIf\s+(.+?)\s+Then$/i.exec(trimmed) ?? /^#ElseIf\s+(.+)$/i.exec(trimmed);
    if (elifMatch) {
      const current = stack[stack.length - 1];
      if (!current) {
        return "";
      }
      if (current.taken) {
        current.active = false;
      } else {
        const result = evaluate(elifMatch[1]);
        current.active = result === true;
        current.taken = current.active;
      }
      return "";
    }
    if (/^#Else$/i.test(trimmed)) {
      const current = stack[stack.length - 1];
      if (current) {
        if (current.taken) {
          current.active = false;
        } else {
          current.active = true;
          current.taken = true;
        }
      }
      return "";
    }
    if (/^#End\s*If$/i.test(trimmed)) {
      stack.pop();
      return "";
    }
    const active = stack.every((item) => item.active);
    if (active) {
      activeLineCount += 1;
      return line;
    }
    inactiveLineCount += 1;
    return "";
  });
  return { processed: output.join("\n"), constants, activeLineCount, inactiveLineCount, unknownConstant };
}
function buildServices() {
  const shared = inject(createDefaultSharedCoreModule(EmptyFileSystem), VbaMiniGeneratedSharedModule);
  const services = inject(createDefaultCoreModule({ shared }), VbaMiniGeneratedModule);
  shared.ServiceRegistry.register(services);
  return services;
}
async function main() {
  const services = buildServices();
  const parse = parseHelper(services);
  const source = [
    "' header comment",
    "#Const DEBUG = 1",
    "Public Sub Test()",
    "#If DEBUG = 1 Then",
    "    Dim x As Long",
    "    x = 1",
    "#Else",
    "    Dim y As Long",
    "    y = This is broken",
    "#End If",
    "    result = result +",
    "End Sub"
  ].join("\n");
  const pre = preprocess(source, {});
  const doc = await parse(pre.processed, { documentUri: URI.parse("file:///cond-1.bas") });
  const root = doc.parseResult.value;
  const procedures = (root.members ?? []).filter(isProcedure);
  const errorLines = doc.parseResult.parserErrors.map((error) => error.token?.startLine);
  const expectedZeroBasedLine = 10;
  const expectedOneBasedLine = expectedZeroBasedLine + 1;
  const originalLines = source.split("\n");
  const processedLines = pre.processed.split("\n");
  const report = {
    measured: true,
    langium: "4.4.0",
    technique: "\u6E05\u7A7A\u975E\u751F\u6548\u5206\u652F\u884C\u5185\u5BB9\u4F46\u4FDD\u7559\u884C\u6570\uFF0C\u4EE4\u751F\u6548\u5206\u652F\u884C\u5217\u4F4D\u7F6E\u4E0E\u539F\u59CB\u6587\u4EF6\u4E00\u81F4",
    constants: pre.constants,
    activeLineCount: pre.activeLineCount,
    inactiveLineCount: pre.inactiveLineCount,
    unknownConstant: pre.unknownConstant,
    procedureCount: procedures.length,
    procedureNames: procedures.map((item) => item.name),
    parserErrors: doc.parseResult.parserErrors.length,
    errorLines,
    errorMessages: doc.parseResult.parserErrors.map((error) => error.message),
    lineBase: "Chevrotain token.startLine \u4E3A 1 \u57FA\uFF1BLSP \u884C\u53F7\u4E3A 0 \u57FA",
    inactiveBranchErrorReported: errorLines.includes(9),
    // 非生效分支行（0 基 8）对应 1 基 9
    activeErrorLineMatchesOriginal: errorLines.includes(expectedOneBasedLine),
    originalLineAtExpected: originalLines[expectedZeroBasedLine],
    processedLineAtExpected: processedLines[expectedZeroBasedLine],
    lineCountPreserved: originalLines.length === processedLines.length
  };
  writeFileSync("poc/langium/cond-result.json", JSON.stringify(report, null, 2) + "\n");
  console.log(JSON.stringify(report, null, 2));
}
main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
