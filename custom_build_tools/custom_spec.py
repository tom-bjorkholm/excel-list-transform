
"""Repository-specific build specification for common_build_tools."""

from typing import Optional
from build_spec import BuildSpec


def custom_spec() -> Optional[BuildSpec]:
    """Return custom build spec for this repository."""
    return BuildSpec(python_layout_max_name_length=25)
