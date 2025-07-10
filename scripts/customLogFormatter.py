import logging

# Adapted from: https://stackoverflow.com/questions/384076/how-can-i-color-python-logging-output
class CustomFormatter(logging.Formatter):

    grey = "\x1b[38;20m"
    green = "\x1b[32;20m"
    yellow = "\x1b[33;20m"
    red = "\x1b[31;20m"
    bold_red = "\x1b[31;1m"
    reset = "\x1b[0m"

    start = "[%(asctime)s] - %(filename)s:%(lineno)04d -"
    format = " [%(levelname)s] " 
    end = "- %(message)s"

    FORMATS = {
        logging.DEBUG: start + grey + format + reset + end,
        logging.INFO: start + green + format + reset + end,
        logging.WARNING: start + yellow + format + reset + end,
        logging.ERROR: start + red + format + reset + end,
        logging.CRITICAL: start + bold_red + format + reset + end
    }

    def format(self, record):
        log_fmt = self.FORMATS.get(record.levelno)
        formatter = logging.Formatter(log_fmt, datefmt="%Y-%m-%dT%H:%M:%S")
        return formatter.format(record)