import win32com.client
import flask
from flask import Flask, render_template, request, session, redirect, url_for
import sqlite3

dcom_obj = win32com.client.Dispatch("DCOM.Server")