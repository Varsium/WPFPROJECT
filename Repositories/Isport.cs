using System.Collections.Generic;
using EindOpdrachtLeandro.model;

namespace EindOpdrachtLeandro.Repositories
{
    interface Isport
    {
        IList<Sport> GetAll();
    }
}
