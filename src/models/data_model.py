from pydantic import BaseModel, Field, ValidationError
from typing import Dict
from datetime import datetime

class SensorData(BaseModel):
    timestamp: datetime
    measurement: str = Field(..., example="sensor_data")
    values: Dict[str, float]

    @classmethod
    def validate_data(cls, data: Dict):
        """
        Validate incoming data.
        """
        try:
            return cls(**data)
        except ValidationError as e:
            print(f"Validation failed: {e.json()}")
            raise