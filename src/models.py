from pydantic import BaseModel

class PricesDemo(BaseModel):
    asofdate: str
    IUSB : float
    IVE : float
    SHV : float
    SPY : float
    VGSLX: float
