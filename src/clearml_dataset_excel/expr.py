from __future__ import annotations

import ast
from dataclasses import dataclass
from typing import Any


class ExprError(ValueError):
    pass


@dataclass(frozen=True)
class _EvalContext:
    namespace: dict[str, Any]


def eval_expr(expr: str, namespace: dict[str, Any]) -> Any:
    """
    Safely evaluate a simple arithmetic expression.

    Allowed:
    - numeric constants
    - column names as identifiers (namespace lookup)
    - +, -, *, / and parentheses
    """
    try:
        tree = ast.parse(expr, mode="eval")
    except SyntaxError as e:
        raise ExprError(f"Invalid expression: {expr}") from e

    return _eval_node(tree.body, _EvalContext(namespace=namespace))


def _eval_node(node: ast.AST, ctx: _EvalContext) -> Any:
    if isinstance(node, ast.Constant):
        if isinstance(node.value, (int, float)) and not isinstance(node.value, bool):
            return node.value
        raise ExprError(f"Unsupported constant: {node.value!r}")

    # Py<=3.7 compatibility (not needed here, but safe)
    if isinstance(node, ast.Num):  # pragma: no cover
        return node.n

    if isinstance(node, ast.Name):
        if node.id not in ctx.namespace:
            raise ExprError(f"Unknown identifier: {node.id}")
        return ctx.namespace[node.id]

    if isinstance(node, ast.UnaryOp):
        operand = _eval_node(node.operand, ctx)
        if isinstance(node.op, ast.UAdd):
            return +operand
        if isinstance(node.op, ast.USub):
            return -operand
        raise ExprError(f"Unsupported unary op: {type(node.op).__name__}")

    if isinstance(node, ast.BinOp):
        left = _eval_node(node.left, ctx)
        right = _eval_node(node.right, ctx)
        if isinstance(node.op, ast.Add):
            return left + right
        if isinstance(node.op, ast.Sub):
            return left - right
        if isinstance(node.op, ast.Mult):
            return left * right
        if isinstance(node.op, ast.Div):
            return left / right
        raise ExprError(f"Unsupported binary op: {type(node.op).__name__}")

    if isinstance(node, ast.Expr):  # pragma: no cover
        return _eval_node(node.value, ctx)

    raise ExprError(f"Unsupported expression node: {type(node).__name__}")

