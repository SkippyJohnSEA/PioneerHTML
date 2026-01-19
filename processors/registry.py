from . import script_president, script_social, script_calendar

SCRIPTS = {
    "Social Events Accordion": {
        "func": script_social.run, # now expects (input_path, original_name)
        "needs_target": False
    },
    "Calendar Accordion": {
        "func": script_calendar.run, # now expects (input_path, original_name)
        "needs_target": False
    },
    "Presidents Table": {
        "func": script_president.run, # now expects (input_path, original_name)
        "needs_target": False
    }
}