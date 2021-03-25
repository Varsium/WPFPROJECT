using System.Collections.Generic;
using EindOpdrachtLeandro.model;

namespace EindOpdrachtLeandro.Repositories
{
    interface Ihobby
    {
        IList<Hobby> GetAll();
    }
}
