#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
CustomLogger.py
~~~~~~~~~~~~

This module implements the Logging package for Python. Based on PEP 282 and comments thereto in
comp.lang.python.

:copyright: (c) 2020 by Chandrayee Kumar.All Rights Reserved.
:license: Ericsson, see LICENSE for more details.

To use, simply 'import CustomLogger' set log = getLogger('root') and log away!
"""

import logging
import logging.handlers
from datetime import datetime
from logging.handlers import RotatingFileHandler
import os
import sys
# sys.path.append('C:\\Users\\ekcuhma\\OneDrive - Ericsson AB\\Local\\design automation\\ims-lld-automation')
import yaml

__author__  = "Chandrayee Kumar <chandrayee.kumar@ericsson.com>"
__status__  = "development"
# The following module attributes are no longer updated.
__version__ = "0.1"
__date__    = "20 April 2020"
#
#resource_dict is used as the dictionary and used for fetching log path
#
# if 'Core' in os.path.realpath("resource.yaml"): 
#   file_path=os.path.realpath("resource.yaml").replace("\Core","")
# else:
#   file_path=os.path.realpath("resource.yaml")  
  # print(file_path)
# resource_dict = yaml.safe_load(open("../../log"))
# print(os.path.abspath(file_path.replace("resource.yaml","")+resource_dict["Resources"]["Log_Path"]))
logPath='../../log'
# print(logPath)


def getLogger(name='root', loglevel='DEBUG'):
  
  """
  The level is one of the predefined levels (CRITICAL, ERROR, WARNING,
  INFO, DEBUG) then you get the corresponding string. If you have
  associated levels with names using addLevelName then the name you have
  associated with 'level' is returned.
      
  RotatingFileHandler for logging to a set of files, which switches from one file
  to the next when the current file reaches a certain size.
    
  """
  logger = logging.getLogger(name)

  # if logger 'name' already exists, return it to avoid logging duplicate
  # messages by attaching multiple handlers of the same type
  if logger.handlers:
    return logger
  # if logger 'name' does not already exist, create it and attach handlers
  else:
    # set logLevel to loglevel or to INFO if requested level is incorrect
    loglevel = getattr(logging, loglevel.upper(), logging.DEBUG)
    logger.setLevel(loglevel)
    fmt = '%(asctime)s %(filename)-18s %(levelname)-8s: %(message)s'
    fmt_date = '%Y-%m-%dT%T'
    formatter = logging.Formatter(fmt, fmt_date)
    fileName = datetime.now().strftime('CIQ_%d-%m-%Y.log')
    handler = logging.handlers.RotatingFileHandler("{0}/{1}".format(logPath, fileName), maxBytes=(1048576*5), backupCount=0,encoding=None)
    handler.setFormatter(formatter)
    logger.addHandler(handler)
    

    if logger.name == 'root':
      logger.warning('****************************CIQ PROCESS START*****************************')
      logger.warning('Running: %s %s',
                     os.path.basename(sys.argv[0]),
                     ' '.join(sys.argv[1:]))
    return logger