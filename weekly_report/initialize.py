import logging


def initialize():
    """  set up logger and if needed erase files, etc  """
    logging.root.handlers = []
    logging.getLogger().setLevel(100)
    logger_a = logging.getLogger('my_logger')
    logger_a_handler = logging.StreamHandler()
    logger_a_handler.setLevel(logging.DEBUG)
    logger_a_formatter = logging.Formatter('%(levelname)s - %(message)s')
    # logger_a_formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s', "%Y-%m-%d %H:%M:%S")
    logger_a_handler.setFormatter(logger_a_formatter)
    logger_a.addHandler(logger_a_handler)

    return logger_a
