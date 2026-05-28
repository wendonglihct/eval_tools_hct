import logging
from logging.handlers import RotatingFileHandler
import os

def setup_bot_logger(log_dir):
    os.makedirs(log_dir, exist_ok=True)

    logger = logging.getLogger("chatbot")
    logger.setLevel(logging.INFO)

    if logger.handlers:
        return logger

    handler = RotatingFileHandler(
        filename=os.path.join(log_dir, "chatbot.log"),
        maxBytes=100 * 1024 * 1024,  # 100MB
        backupCount=5,
        encoding="utf-8",
    )

    formatter = logging.Formatter(
        "[%(asctime)s] [%(levelname)s] "
        "[%(threadName)s] "
        "%(filename)s:%(lineno)d - %(message)s"
    )
    handler.setFormatter(formatter)

    logger.addHandler(handler)
    logger.propagate = False
    return logger