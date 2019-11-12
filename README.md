# Multivalued Mail Merge for Microsoft Word in VBA

VBA code to import into a `.docm` document so that mail-marge handle groups of multivalued field.

## Example

Take a CSV file (or anything usable in Word mail-merge) with new-line separated values :

| Story              | Characters_FirstName                      | Characters_Surname                      |
|--------------------|-------------------------------------------|-----------------------------------------|
| Madame Bovary      | Emma<br>Charles<br>Rodolphe<br>Leon       | Bovary<br>Bovary<br>Boulanger<br>Dupuis |
| Matter of Britain  | Arthur<br>Merlin<br>Guinevere<br>Lancelot | Pendragon<br>Inchanter<br><br>Du Lac    |

### Take a `.docm` document with mail-merge fields

The story `«Story»` comprises the following characters :

| First Name               | Surname                | 
|--------------------------|------------------------|
| `«Characters_FirstName»` | `«Characters_Surname»` |

### When the user launches Mail Merge, the standard result is as follow :
---------------

The story Madame Bovary comprises the following characters :

| First Name                          | Surname                                 | 
|-------------------------------------|-----------------------------------------|
| Emma<br>Charles<br>Rodolphe<br>Leon | Bovary<br>Bovary<br>Boulanger<br>Dupuis |

---------------

The story Matter of Britain comprises the following characters :

| First Name                                      | Surname                              | 
|-------------------------------------------------|--------------------------------------|
| Arthur<br>Merlin<br>Guinevere<br>Lancelot       | Pendragon<br>Inchanter<br><br>Du Lac |


### Using this set of VBA macros, the result is as follow :
---------------

The story Madame Bovary comprises the following characters :

| First Name | Surname   | 
|------------|-----------|
| Emma       | Bovary    |
| Charles    | Bovary    |
| Rodolphe   | Boulanger |
| Leon       | Dupuis    |

---------------

The story Matter of Britain comprises the following characters :

| First Name | Surname   | 
|------------|-----------|
| Arthur     | Pendragon |
| Merlin     | Inchanter |
| Guinevere  |           |
| Lancelot   | Du Lac    |
