class ResolvedIncident:
    def __init__(self, title, created_at, resolved_time):
        self.title = title
        self.created_at = created_at
        self.resolved_time = resolved_time

    def __eq__(self, other):
        return self.title==other.title\
            and self.created_at==other.created_at\
            and self.resolved_time==other.resolved_time

    def __hash__(self):
        return hash(('title',self.title,
                    'created_at',self.created_at,
                    'resolved_time',self.resolved_time))

class OpenIncident:
    def __init__(self, id, title, created_at):
        self.id = id
        self.title = title
        self.created_at = created_at

    def __eq__(self, other):
        return self.id==other.id\
            and self.title==other.title\
            and self.created_at==other.created_at

    def __hash__(self):
        return hash(('id',self.id,
                    'title',self.title,
                    'created_at',self.created_at))
