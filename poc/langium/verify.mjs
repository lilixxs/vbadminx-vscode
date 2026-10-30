// poc/langium/src/verify.ts
import { readFileSync, writeFileSync } from "node:fs";
import { EmptyFileSystem } from "langium";
import { createDefaultCoreModule, createDefaultSharedCoreModule, inject } from "langium";
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
function isCallStatement(item) {
  return reflection.isInstance(item, CallStatement.$type);
}
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
function isFunctionCall(item) {
  return reflection.isInstance(item, FunctionCall.$type);
}
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

// poc/langium/src/verify.ts
function buildServices() {
  const shared = inject(createDefaultSharedCoreModule(EmptyFileSystem), VbaMiniGeneratedSharedModule);
  const services2 = inject(createDefaultCoreModule({ shared }), VbaMiniGeneratedModule);
  shared.ServiceRegistry.register(services2);
  return services2;
}
var services = buildServices();
var parse = parseHelper(services);
function proceduresOf(root) {
  return (root.members ?? []).filter(isProcedure);
}
var scenarioIndex = 0;
async function scenario(name, purpose, source, expectedProcedures) {
  scenarioIndex += 1;
  const doc = await parse(source, { documentUri: `file:///poc-${scenarioIndex}.bas` });
  const root = doc.parseResult.value;
  const procedures = proceduresOf(root);
  const names = procedures.map((item) => item.name);
  let resolvedCalls = 0;
  let unresolvedCalls = 0;
  const countRef = (ref) => {
    if (ref) {
      resolvedCalls += 1;
    } else {
      unresolvedCalls += 1;
    }
  };
  const seen = /* @__PURE__ */ new Set();
  const walk = (node) => {
    if (!node || typeof node !== "object" || seen.has(node)) {
      return;
    }
    seen.add(node);
    if (isCallStatement(node)) {
      countRef(node.callee?.ref);
    } else if (isFunctionCall(node)) {
      countRef(node.callee?.ref);
    }
    for (const key of Object.keys(node)) {
      if (key.startsWith("$")) {
        continue;
      }
      const value = node[key];
      if (Array.isArray(value)) {
        for (const item of value) {
          walk(item);
        }
      } else if (value && typeof value === "object") {
        walk(value);
      }
    }
  };
  walk(root);
  const found = new Set(names);
  const missingProcedures = expectedProcedures.filter((item) => !found.has(item));
  return {
    name,
    purpose,
    parserErrors: doc.parseResult.parserErrors.length,
    procedureCount: procedures.length,
    procedureNames: names,
    expectedProcedures,
    missingProcedures,
    recovered: missingProcedures.length === 0,
    resolvedCalls,
    unresolvedCalls,
    firstErrors: doc.parseResult.parserErrors.slice(0, 5).map((error) => error.message)
  };
}
var computeTail = `Public Function Compute(a As Long) As Long
    Compute = a
End Function
`;
async function main() {
  const results = [];
  results.push(await scenario(
    "A-correct-baseline",
    "\u57FA\u7EBF\uFF1A\u4E09\u4E2A\u5B8C\u6574\u8FC7\u7A0B\uFF0C\u9A8C\u8BC1\u89E3\u6790\u4E0E\u4EA4\u53C9\u5F15\u7528\u53EF\u7528",
    `Public Sub First()
    Dim x As Long
    x = 1
End Sub

Public Function Second() As Long
    Second = Compute(1)
End Function

${computeTail}`,
    ["First", "Second", "Compute"]
  ));
  results.push(await scenario(
    "B-missing-End-Sub",
    "\u5BB9\u9519\u6838\u5FC3\uFF1A\u9996\u8FC7\u7A0B\u7F3A End Sub\uFF0C\u68C0\u67E5\u540E\u7EED\u8FC7\u7A0B\u662F\u5426\u4ECD\u8FDB\u5165 AST",
    `Public Sub First()
    Dim x As Long
    x = 1

Public Function Second() As Long
    Second = Compute(1)
End Function

${computeTail}`,
    ["Second", "Compute"]
  ));
  results.push(await scenario(
    "C-statement-error",
    "\u8BED\u53E5\u7EA7\u9519\u8BEF\uFF1A\u8FC7\u7A0B\u5185 If \u7F3A Then\uFF0C\u68C0\u67E5\u540E\u7EED\u8FC7\u7A0B\u662F\u5426\u4ECD\u8FDB\u5165 AST",
    `Public Sub First()
    If x > 1
        x = 2
    End If
End Sub

Public Function Second() As Long
    Second = Compute(1)
End Function

${computeTail}`,
    ["Second", "Compute"]
  ));
  const sample = "D:/SVN_CQ24/trunk/doc/process/BAS(\u8282\u80FD)/02-\u751F\u6210\u7A0B\u5E8F&\u751F\u6210\u7ED3\u679C\u5907\u4EFD/02-\u751F\u6210\u7A0B\u5E8F/X89-\u8F85\u52A9\u5C0F\u5DE5\u5177/vba_sync/X89-\u8F85\u52A9\u5C0F\u5DE5\u5177-V1_4.xlsm/\u6A21\u5757/CQ24_AdvancedSort.bas";
  let realSample = { skipped: "\u6837\u4F8B\u4E0D\u53EF\u8BFB" };
  try {
    const text = readFileSync(sample, "utf8");
    const doc = await parse(text, { documentUri: "file:///CQ24_AdvancedSort.bas" });
    const root = doc.parseResult.value;
    const procedures = proceduresOf(root);
    realSample = {
      note: "\u4EC5\u4F9B\u53C2\u8003\uFF1Amini \u8BED\u6CD5\u4EC5\u8986\u76D6\u8FC7\u7A0B/\u5757/\u8868\u8FBE\u5F0F\uFF0C\u672A\u8986\u76D6 Type \u4E4B\u5916\u7684 Enum/Declare/On Error/GoTo/\u884C\u6807\u7B7E/Select Case/Exit/\u6761\u4EF6\u7F16\u8BD1\u7B49\uFF1B\u9519\u8BEF\u6570\u4E0E\u8FC7\u7A0B\u6570\u4E0D\u4EE3\u8868 Langium \u80FD\u529B",
      file: "CQ24_AdvancedSort.bas",
      parserErrors: doc.parseResult.parserErrors.length,
      procedureCount: procedures.length,
      procedureNames: procedures.slice(0, 10).map((item) => item.name),
      // 只输出错误消息与行号，不输出业务源码内容。
      firstErrors: doc.parseResult.parserErrors.slice(0, 6).map((error) => ({
        message: error.message,
        line: error.token?.startLine,
        offset: error.token?.startOffset
      }))
    };
  } catch (error) {
    realSample = { error: error instanceof Error ? error.message : "\u8BFB\u53D6\u6216\u89E3\u6790\u5931\u8D25" };
  }
  const report = { measured: true, langium: "4.4.0", generated: "langium-cli 4.4.0", scenarios: results, realSample };
  writeFileSync("poc/langium/poc-result.json", JSON.stringify(report, null, 2) + "\n");
  console.log(JSON.stringify(report, null, 2));
}
main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
