/*
* DropDownTree client side object model
*/
function DropDownTree(id,option) {
    this.Id = id;
    this.pressedImage = option["pressed"];
    this.normalImage = option["normal"];    
}

DropDownTree.prototype.isExpand = false;

/*
* Open or close the tree panel
* Author: Yin.P
*/
DropDownTree.prototype.switchDropdownTree = function() {
    var objDropdownButton = document.getElementById(this.Id + "_dropdownbutton");
    var objTreeFrame = document.getElementById(this.Id + '_tree_frame');
    if (this.isExpand == false) {
        objTreeFrame.style.display = "block";
        objTreeFrame.style.zIndex = 9999;
        objDropdownButton.src = this.pressedImage;
        this.isExpand = true;
    }
    else {
        objTreeFrame.style.display = "none";
        objDropdownButton.src = this.normalImage;
        this.isExpand = false;
    }

    window.event.cancelBubble = true;
};

/*
* Close the tree panel
* Author: Yin.P
*/
DropDownTree.prototype.collapseDropdownTree = function() {
    var objDropdownButton = document.getElementById(this.Id + "_dropdownbutton");
    var objTreeFrame = document.getElementById(this.Id + '_tree_frame');
    if (this.isExpand) {
        objTreeFrame.style.display = "none";
        objDropdownButton.src = this.normalImage;
        this.isExpand = false;
    }
};

/*
* Set text and value
* Author: Yin.P
*/
DropDownTree.prototype.selectNode = function(id, text, value, dataitemString) {
    var txt = document.getElementById(id + "__SelectedText");
    var val = document.getElementById(id + "__SelectedValue");
    var dataitem = document.getElementById(this.Id + "__SelectedDataItem");

    if (txt != null) {
        txt.value = text;
    }

    if (val != null) {
        val.value = value;
    }

    if (dataitem != null) {
        dataitem.value = dataitemString;
    }

    this.collapseDropdownTree();
};

/*
* Client Side Interface
* Get selected text
* Author: Yin.P
*/
DropDownTree.prototype.getSelectedText = function() {
    var txt = document.getElementById(this.Id + "__SelectedText");
    return txt.value;
};

/*
* Client Side Interface
* Get selected value
* Author: Yin.P
*/
DropDownTree.prototype.getSelectedValue = function() {
    var val = document.getElementById(this.Id + "__SelectedValue");
    return val.value;
};

/*
* Client Side Interface
* Get selected dataitem. it contains all the columns in datasource.
* This method return a JSON object
* Author: Yin.P
*/
DropDownTree.prototype.getSelectedDataItem = function() {
    var dataitem = document.getElementById(this.Id + "__SelectedDataItem");
    return eval("(" + dataitem.value + ")");
};

/*
* Author: Yin.P
*/
DropDownTree.prototype.Init = function() {
    var bodyObj = document.getElementsByTagName("body")[0];
    var oldBodyClick = bodyObj.onclick;
    if (typeof bodyObj.onclick != 'function') {
        bodyObj.onclick = this.collapseDropdownTree;
    } else {
        bodyObj.onclick = function() {
            oldBodyClick();
            this.collapseDropdownTree();
        }
    }
};