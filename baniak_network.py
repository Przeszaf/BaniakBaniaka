from pyvis.network import Network

class Connection:
    firstPerson = None
    secondPerson = None
    value = 0

    def contains(self, person1, person2):
        return self.firstPerson is person1 and self.secondPerson is person2 or self.firstPerson is person2 and self.secondPerson is person1


def createNetwork(sessions, people):
    connections = []
    for session in sessions:
        for i in range(len(session.people)):
            for j in range(i + 1, len(session.people)):
                index = None
                for connectionIndex in range(len(connections)):
                    if connections[connectionIndex].contains(session.people[i], session.people[j]):
                        index = connectionIndex
                        break
                if index is None:
                    newConnection = Connection()
                    newConnection.firstPerson = session.people[i]
                    newConnection.secondPerson = session.people[j]
                    newConnection.value += 1
                    connections.append(newConnection)
                else:
                    connections[connectionIndex].value += 1


    net = Network(width='100%', height = '100%', bgcolor="#000000", font_color="#FFFFFF")


    for person in people:
        net.add_node(n_id=person.displayName, color=person.color(), physics=False, x=person.x, y=person.y)

    for connection in connections:
        color = getCombinedColor(connection.firstPerson.color(), connection.secondPerson.color())
        net.add_edge(connection.firstPerson.displayName, connection.secondPerson.displayName, value=connection.value, color=color)

    net.show('bcu_connections.html')

def getCombinedColor(color1, color2):
    firstColor = tuple(int(color1.lstrip('#')[i:i+2], 16) for i in (0, 2, 4))
    secondColor = tuple(int(color2.lstrip('#')[i:i+2], 16) for i in (0, 2, 4))
    combinedColor = (int((firstColor[0] + secondColor[0]) / 2), int((firstColor[1] + secondColor[1]) / 2), int((firstColor[2] + secondColor[2]) / 2))
    return '#%02x%02x%02x' % combinedColor