#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sat Jun 27 20:57:17 2020

@author: dsp
"""
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time

# Stating the location of Chrome webdriver to allow access to selenium
location = '/home/dsp/Downloads/chromedriver_linux64/chromedriver'
driver = webdriver.Chrome(location)

# main function to send whatsapp message

def send_message(name, message):
    # Searching for the person to send the message to
    new_msg = "/html/body/div[1]/div/div/div[3]/div/div[1]/div/label/div/div[2]"
    new = driver.find_element_by_xpath(new_msg)
    new.click()
    new.send_keys(name + Keys.ENTER)
    
    # Finding the text field and clicking on it
    inp_xpath = '//*[@id="main"]/footer/div[1]/div[2]/div/div[2]'
    input_box = driver.find_element_by_xpath(inp_xpath)
    input_box.click()
    
    # typing in the message and sending it
    input_box.send_keys(message + Keys.ENTER)
    time.sleep(1)


def form_string(string):
    # "_" at the beginning and at the end to get the final message
    # in italics
    
    msg = '_You have desk duties from '
    for i in string:
        msg = msg + i
    msg = msg + '_'
    return msg


# Initiation function to start-up web.whatsapp.com and sign-in         
def initiate():
    driver.get('https://web.whatsapp.com/')
    input('Press Enter key after scanning QR Code...')

    
# Function to send introduction messages to everyone
def send_intro_message(name_list):
    message = "_"
    msg = ["Hi, really sorry to bother you at this time. ",
               "Just a gentle reminder regarding your desk duties ",
               "for the upcoming week_"]
    for i in msg:
        message = message + i
        
    for i in name_list:
        send_message(i, message)


# function to send whatsapp message regarding slots
   
def send_slots_message(name, day, slot, place):
    # Setting up the message
    string = [slot , ' at ' , place , ' on ', day]
    message = form_string(string)
    send_message(name, message)
    
    

