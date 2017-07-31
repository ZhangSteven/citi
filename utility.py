# coding=utf-8
# 
# from config_logging package, provides a config object (from config file)
# and a logger object (logging to a file).
# 

import configparser, os
from config_logging.file_logger import get_file_logger



class InvalidDatamode(Exception):
	pass



def get_current_directory():
	"""
	Get the absolute path to the directory where this module is in.

	This piece of code comes from:

	http://stackoverflow.com/questions/3430372/how-to-get-full-path-of-current-files-directory-in-python
	"""
	return os.path.dirname(os.path.abspath(__file__))



def _load_config():
	"""
	Read the config file, convert it to a config object.
	"""
	cfg = configparser.ConfigParser()
	cfg.read(os.path.join(get_current_directory(), 'citi.config'))
	return cfg



# initialized only once when this module is first imported by others
if not 'config' in globals():
	config = _load_config()



# def get_log_directory():
# 	"""
# 	The directory where the log file resides.
# 	"""
# 	global config
# 	directory = config['logging']['directory']
# 	if directory == '':
# 		directory = get_current_path()

# 	return directory



def _setup_logging():
	global config
	directory = config['logging']['directory']
	if directory == '':
		directory = get_current_directory()
    fn = os.path.join(directory, config['logging']['log_file'])
    log_level = config['logging']['log_level']
    return get_file_logger(fn, log_level)



# initialized only once when this module is first imported by others
if not 'logger' in globals():
	logger = _setup_logging()



def get_datemode():
	"""
	Read datemode from the config object and return it (in integer)
	"""
	global config
	d = config['excel']['datemode']
	try:
		datemode = int(d)
	except:
		logger.error('get_datemode(): invalid datemode value: {0}'.format(d))
		raise InvalidDatamode()

	return datemode



def get_input_directory():
	"""
	Where the input files reside.
	"""
	global config
	directory = config['input']['directory']
	if directory == '':
		directory = get_current_directory()

	return directory