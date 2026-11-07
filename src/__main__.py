import asyncio
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from main import main  # noqa: E402

asyncio.run(main())
