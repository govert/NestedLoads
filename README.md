# NestedLoads

Test nested loading of add-ins

* Main loads a ribbon COM add-in
* The ribbon from Main implements OnConnection which loads Child1 add-in
* Child1 implements AutoOpen and loads Child2 add-in
* Child2 loads a ribbon COM add-in

Debug output:

```
Main AutoOpen
Main Ribbon OnConnection - Enter
Child1 AutoOpen - Enter
Child2 AutoOpen
Child2 Ribbon OnConnection - Enter
Child1 AutoOpen - Exit
Main Ribbon OnConnection - Exit
```

