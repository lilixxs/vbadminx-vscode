grammar VbaMini

entry Module:
    (members+=ModuleMember | NEWLINE)*;

ModuleMember:
    Procedure | ModuleDeclaration | OptionStatement | TypeDeclaration;

OptionStatement:
    'Option' name=ID (value=(ID | NUMBER))?;

TypeDeclaration:
    ('Public' | 'Private')? 'Type' name=ID NEWLINE
        (fields+=FieldDeclaration)+
    'End' 'Type';

FieldDeclaration:
    name=ID ('As' type=ID)? NEWLINE;

ModuleDeclaration:
    ('Public' | 'Private')? 'Dim' name=ID ('As' type=ID)?;

Procedure:
    ('Public' | 'Private')? procType=('Sub' | 'Function') name=ID
    '(' (params+=Parameter (',' params+=Parameter)*)? ')'
    (('As') returnType=ID)?
    NEWLINE
        (statements+=Statement)*
    'End' endType=('Sub' | 'Function');

Parameter:
    ('Optional')? ('ByVal' | 'ByRef')? name=ID ('As' type=ID)?;

Statement:
      Declaration
    | Assignment
    | CallStatement
    | IfStatement
    | ForStatement
    | WithStatement
    | DoStatement
    | WhileStatement
    | NEWLINE;

Declaration:
    ('Dim' | 'Static') name=ID ('As' type=ID)?;

// VBA 的赋值与函数返回语句语法前缀相同（都形如 `Name = expr`），LL(k) 无法区分。
// 本 PoC 合并为一条规则；真实实现需靠“是否在同名 Function 内”的语义判断区分。
Assignment:
    ('Let' | 'Set')? target=QualifiedName '=' value=Expression;

CallStatement:
    'Call'? callee=[Procedure:QualifiedName] '(' (args+=Expression (',' args+=Expression)*)? ')';

IfStatement:
    'If' condition=Expression 'Then' NEWLINE
        (thenBlock+=Statement)*
    ('ElseIf' elifConditions+=Expression 'Then' NEWLINE (elifBlocks+=Statement)*)*
    ('Else' NEWLINE (elseBlock+=Statement)*)?
    'End' 'If';

ForStatement:
    'For' var=ID '=' from=Expression 'To' to=Expression NEWLINE
        (body+=Statement)*
    'Next' nextVar=ID?;

WhileStatement:
    'While' condition=Expression NEWLINE
        (body+=Statement)*
    'Wend';

DoStatement:
    'Do' ('While' | 'Until')? condition=Expression? NEWLINE
        (body+=Statement)*
    'Loop';

WithStatement:
    'With' target=QualifiedName NEWLINE
        (body+=Statement)*
    'End' 'With';

Expression:
    Literal
    | Reference
    | BinaryExpression
    | UnaryExpression
    | ParenthesizedExpression;

BinaryExpression infers Expression:
    left=PrimaryExpression (operators+=BinaryOperator operands+=PrimaryExpression)*;

BinaryOperator returns string:
    '+' | '-' | '*' | '/' | '&' | '=' | '<' | '>' | '<=' | '>=' | '<>' | 'And' | 'Or';

UnaryExpression infers Expression:
    operator=('-' | 'Not') operand=PrimaryExpression;

ParenthesizedExpression infers Expression:
    '(' inner=Expression ')';

PrimaryExpression infers Expression:
    Literal | FunctionCall | Reference | ParenthesizedExpression;

FunctionCall:
    callee=[Procedure:QualifiedName] '(' (args+=Expression (',' args+=Expression)*)? ')';

Literal:
    value=(STRING | NUMBER | BOOLEAN);

Reference:
    target=QualifiedName;

QualifiedName returns string:
    ID ('.' ID)*;

terminal ID: /[_a-zA-Z][\w_]*[\$\%\&\@\^]?/;
terminal STRING: /"([^"]|"")*"/;
terminal NUMBER returns number: /[0-9]+(\.[0-9]+)?/;
terminal BOOLEAN returns boolean: /true|false/;
hidden terminal WS: /[ \t]+/;
terminal NEWLINE: /[\r\n]+/;
hidden terminal SL_COMMENT: /'[^\r\n]*/;
