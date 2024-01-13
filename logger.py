import logging

# Set up logging configuration
logging.basicConfig(level=logging.INFO)  # Default to INFO level

def logger_decorator(func):
    def wrapper(*args, **kwargs):
        logger = logging.getLogger(func.__name__)

        # Log function entry
        logger.info(f"Entering {func.__name__} with args: {args}, kwargs: {kwargs}")

        result = func(*args, **kwargs)

        # Log function exit
        logger.info(f"Exiting {func.__name__} with result: {result}")

        return result
    return wrapper