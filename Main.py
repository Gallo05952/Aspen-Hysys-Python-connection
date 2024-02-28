import win32com.client

simulation_path = r'file_path.hsc'

def open_hysys_simulation(simulation_path):
    hysys_app = win32com.client.Dispatch("HYSYS.Application")
    hysys_sim = hysys_app.SimulationCases.Open(simulation_path)
    hysys_app.ChangePreferencesToMinimizePopupWindows(True) # This is to avoid popups
    Solver = hysys_sim.Solver
    Solver.CanSolve = False #This set the solver on hold on
    Blocks = hysys_sim.Flowsheet.Operations
    Streams = hysys_sim.Flowsheet.Streams
    return hysys_app, hysys_sim, Blocks, Streams

def close_hysys_simulation(hysys_app, hysys_sim):
    hysys_sim.Close()
    hysys_app.Quit()

def consumiCompressore(spreadsheet): # this version doen't use the global variables, in this way if u use the same template for each excel u can reuse this function
    cella=spreadsheet.Cell(0,6)
    W=cella.CellValue
    cella=spreadsheet.Cell(0,7)
    Q=cella.CellValue
    return W,Q

hysys_app, hysys_sim, Blocks, Streams = open_hysys_simulation(simulation_path)
excel=Blocks.Item("consumi") #this is the name of my spreadsheet. It is in italian word to consumptions
Work,Duty = consumiCompressore(excel)
print("Compressor Work:", Work)
print("Compressor Duty:", Duty)
close_hysys_simulation(hysys_app, hysys_sim)

