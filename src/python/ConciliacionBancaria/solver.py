import random
from deap import base, creator, tools, algorithms

creator.create("FitnessMin", base.Fitness, weights=(-1.0,))
creator.create("Individual", list, fitness=creator.FitnessMin)

def eval_combinacion(individual, numeros, objetivo):
    suma = sum(n * i for n, i in zip(numeros, individual))
    return abs(objetivo - suma),

def crear_individuo(n):
    return creator.Individual(random.choices([0, 1], k=n))

def configurar_toolbox(numeros, objetivo):
    toolbox = base.Toolbox()
    toolbox.register("individual", crear_individuo, n=len(numeros))
    toolbox.register("population", tools.initRepeat, list, toolbox.individual)
    toolbox.register("evaluate", eval_combinacion, numeros=numeros, objetivo=objetivo)
    toolbox.register("mate", tools.cxTwoPoint)
    toolbox.register("mutate", tools.mutFlipBit, indpb=0.05)
    toolbox.register("select", tools.selTournament, tournsize=3)
    return toolbox

def ejecutar_algoritmo(toolbox, n_generaciones=100, poblacion=500):
    population = toolbox.population(n=poblacion)
    algorithms.eaSimple(population, toolbox, cxpb=0.5, mutpb=0.2, ngen=n_generaciones, verbose=False)
    best_individual = tools.selBest(population, 1)[0]
    return best_individual

def obtener_mejor_combinacion(numeros, objetivo, n_iteraciones=100, valor_minimo = 10):
    mejor_combinacion = None
    mejor_diferencia = float('inf')

    toolbox = configurar_toolbox(numeros, objetivo)

    for _ in range(n_iteraciones):
        best_individual = ejecutar_algoritmo(toolbox)
        suma = sum(n for n, i in zip(numeros, best_individual) if i == 1)
        diferencia = abs(objetivo - suma)
        
        if diferencia < mejor_diferencia:
            mejor_diferencia = diferencia
            mejor_combinacion = [numeros[i] for i in range(len(numeros)) if best_individual[i] == 1]

        if mejor_diferencia == 0 or mejor_diferencia <= valor_minimo:
            break

    return mejor_combinacion, mejor_diferencia