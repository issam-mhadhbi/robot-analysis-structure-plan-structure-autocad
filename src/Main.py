import json

import win32com.client
from win32com.client import CastTo

from Grid import GridPipeline
from Columns import ColumnsPipeline 
from Beams import BeamsPipeline
from Slabs import SlabsPipeline
if __name__ == "__main__":
    robot = win32com.client.Dispatch("Robot.Application")
    robot.Visible = True
    locator = GridPipeline(robot)
    ColumnsPipeline(robot , locator )
    BeamsPipeline(robot, locator)
    SlabsPipeline(robot, locator)
