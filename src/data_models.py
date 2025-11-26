"""
Data models for Z21 file structures.
"""

from dataclasses import dataclass, field
from typing import List, Optional, Dict, Any


@dataclass
class FunctionInfo:
    """Represents function information for a locomotive."""
    function_number: int = 0
    image_name: str = ""  # Icon/image name
    shortcut: str = ""
    position: int = 0
    time: str = "0"
    button_type: int = 0  # 0=switch, 1=push-button, 2=time button
    is_active: bool = True
    
    def __repr__(self):
        return f"F{self.function_number}({self.image_name})"
    
    def button_type_name(self) -> str:
        """Get human-readable button type name."""
        button_types = {
            0: "switch",
            1: "push-button",
            2: "time button",
        }
        return button_types.get(self.button_type, f"type_{self.button_type}")


@dataclass
class Locomotive:
    """Represents a locomotive in the Z21 system."""
    address: int = 0
    name: str = ""
    functions: Dict[int, bool] = field(default_factory=dict)  # Legacy: function number -> active state
    function_details: Dict[int, FunctionInfo] = field(default_factory=dict)  # function number -> FunctionInfo
    cvs: Dict[int, int] = field(default_factory=dict)
    speed: int = 0
    direction: bool = True  # True = forward
    
    # Additional fields from SQLite database
    image_name: str = ""  # Image filename for locomotive
    full_name: str = ""  # Full descriptive name
    railway: str = ""  # Railway company/organization
    description: str = ""  # Detailed description
    article_number: str = ""  # Product/article number
    decoder_type: str = ""  # Decoder type (e.g., "NEM 652")
    build_year: str = ""  # Build/manufacturing year
    buffer_length: str = ""  # Buffer length
    model_buffer_length: str = ""  # Model buffer length
    service_weight: str = ""  # Service weight
    model_weight: str = ""  # Model weight
    rmin: str = ""  # Minimum radius
    ip: str = ""  # IP address (if applicable)
    drivers_cab: str = ""  # Driver's cab identifier
    active: bool = True  # Whether locomotive is active
    speed_display: int = 0  # Speed display setting: 0=km/h, 1=Regulation Step, 2=mph
    categories: List[str] = field(default_factory=list)  # Category names
    crane: bool = False  # Crane function flag
    in_stock_since: str = ""  # Date when locomotive was added to stock
    regulation_step: int = 0  # Regulation step from traction_list
    rail_vehicle_type: int = 0  # Rail vehicle type: 0=Loco, 1=Wagon, 2=Accessory
    
    def __repr__(self):
        return f"Locomotive(address={self.address}, name='{self.name}')"


@dataclass
class Accessory:
    """Represents an accessory (turnout, signal, light) in the Z21 system."""
    address: int = 0
    name: str = ""
    accessory_type: str = "unknown"  # turnout, signal, light, etc.
    state: int = 0
    
    def __repr__(self):
        return f"Accessory(address={self.address}, name='{self.name}', type={self.accessory_type})"


@dataclass
class Layout:
    """Represents track layout configuration."""
    name: str = ""
    track_type: str = ""
    blocks: List[Dict[str, Any]] = field(default_factory=list)
    
    def __repr__(self):
        return f"Layout(name='{self.name}')"


@dataclass
class Settings:
    """Represents system settings."""
    auto_stop: bool = False
    boost_mode: bool = False
    # Add more settings as discovered
    
    def __repr__(self):
        return f"Settings()"


@dataclass
class UnknownBlock:
    """Represents an unknown/unidentified data block."""
    offset: int = 0
    length: int = 0
    data: bytes = b''
    
    def __repr__(self):
        return f"UnknownBlock(offset={self.offset}, length={self.length})"


@dataclass
class Z21File:
    """Root container for Z21 file data."""
    version: Optional[int] = None
    locomotives: List[Locomotive] = field(default_factory=list)
    accessories: List[Accessory] = field(default_factory=list)
    layouts: List[Layout] = field(default_factory=list)
    settings: Optional[Settings] = None
    unknown_blocks: List[UnknownBlock] = field(default_factory=list)
    
    def __repr__(self):
        return (f"Z21File(version={self.version}, "
                f"locomotives={len(self.locomotives)}, "
                f"accessories={len(self.accessories)}, "
                f"layouts={len(self.layouts)}, "
                f"unknown_blocks={len(self.unknown_blocks)})")

