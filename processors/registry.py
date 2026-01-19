from . import script_social, script_calendar, script_presidents

SCRIPTS = {
    "Social Events Accordion": {
        "func": script_social.run,
        "needs_target": False
    },
    "Calendar Accordion": {
        "func": script_calendar.run,
        "needs_target": False
    },
    "Presidents Table": {
        "func": script_presidents.run,
        "needs_target": False
    }
}