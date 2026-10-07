import ast
from dataclasses import dataclass, field
from typing import List, Optional


@dataclass
class CodeElement:
    kind: str
    name: str
    docstring: Optional[str] = None
    source: str = ""
    line_start: int = 0
    line_end: int = 0
    children: List['CodeElement'] = field(default_factory=list)
    decorators: List[str] = field(default_factory=list)
    args: str = ""


def parse_python_code(source: str, module_name: str = "module") -> CodeElement:
    try:
        tree = ast.parse(source)
    except SyntaxError as e:
        raise ValueError(f"Ошибка синтаксиса: {e}")

    lines = source.splitlines()

    def _get_source(node) -> str:
        start = node.lineno - 1
        end = getattr(node, "end_lineno", node.lineno)
        return "\n".join(lines[start:end])

    def _get_decorators(node) -> List[str]:
        return [ast.unparse(d) for d in getattr(node, "decorator_list", [])]

    def _get_args(node) -> str:
        return ast.unparse(node.args) if hasattr(node, "args") else ""

    root = CodeElement(
        kind="module", name=module_name,
        docstring=ast.get_docstring(tree),
        source=source, line_start=1, line_end=len(lines),
    )

    for node in tree.body:
        if isinstance(node, ast.ClassDef):
            cls = CodeElement(
                kind="class", name=node.name,
                docstring=ast.get_docstring(node),
                source=_get_source(node),
                line_start=node.lineno,
                line_end=getattr(node, "end_lineno", node.lineno),
                decorators=_get_decorators(node),
            )
            for item in node.body:
                if isinstance(item, (ast.FunctionDef, ast.AsyncFunctionDef)):
                    cls.children.append(CodeElement(
                        kind="method", name=item.name,
                        docstring=ast.get_docstring(item),
                        source=_get_source(item),
                        line_start=item.lineno,
                        line_end=getattr(item, "end_lineno", item.lineno),
                        decorators=_get_decorators(item),
                        args=_get_args(item),
                    ))
            root.children.append(cls)
        elif isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            root.children.append(CodeElement(
                kind="function", name=node.name,
                docstring=ast.get_docstring(node),
                source=_get_source(node),
                line_start=node.lineno,
                line_end=getattr(node, "end_lineno", node.lineno),
                decorators=_get_decorators(node),
                args=_get_args(node),
            ))
    return root


def collect_leaves(root: CodeElement) -> List[CodeElement]:
    result = []
    for child in root.children:
        result.append(child)
        result.extend(child.children)
    return result
