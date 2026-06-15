"""Точка входа для хостинга (bothost и др.)."""
from bot import main
import asyncio

if __name__ == "__main__":
    asyncio.run(main())
