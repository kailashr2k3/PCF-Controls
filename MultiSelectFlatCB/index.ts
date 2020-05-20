import {IInputs, IOutputs} from "./generated/ManifestTypes";

export class MutiSelectFlatCB implements ComponentFramework.StandardControl<IInputs, IOutputs> {
	private _context : ComponentFramework.Context<IInputs>;
	private _container : HTMLDivElement;
	private _mainContainer : HTMLDivElement;
	private _unorderedList : HTMLUListElement;
	private _errorLabel : HTMLLabelElement;
	public _guidList : string[];
	private _entityName : string;
	private _fieldName : string;
	private _checkBoxChanged : EventListenerOrEventListenerObject;
	private _notifyOutputChanged: () => void;

	// Event listener for changes in the credit card number
 	//private _creditCardNumberChanged: EventListenerOrEventListenerObject;
	/**
	 * Empty constructor.
	 */
	constructor()
	{

	}

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
	 */
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement)
	{
		// Add control initialization code
		this._context = context;
		this._container = container;
		
		this._mainContainer = document.createElement("div");
		this._unorderedList = document.createElement("ul");
		this._errorLabel = document.createElement("label");
		this._unorderedList.classList.add("ks-cboxtags");
		this._mainContainer.classList.add("multiselect-container");
		var messageString = "";
		
		this._notifyOutputChanged = notifyOutputChanged;
		this._checkBoxChanged = this.checkBoxChanged.bind(this);

		if(this._context.parameters.entityNameToTransform.raw != null && this._context.parameters.entityNameToTransform.raw != ""){
			this._entityName = this._context.parameters.entityNameToTransform.raw;
		}

		if(this._context.parameters.fieldName.raw != null && this._context.parameters.fieldName.raw != ""){
			this._fieldName = this._context.parameters.fieldName.raw;
		}
		
		var self = this;
		var filter = "?$select=" + this._fieldName + "," + this._entityName + "id";
		this._context.webAPI.retrieveMultipleRecords(this._entityName, filter).then(
			function success(result) {
				for (var i = 0; i < result.entities.length; i++) {
					var newChkBox = document.createElement("input");
					var newLabel = document.createElement("label");
					var newUList = document.createElement("li");
					
					newChkBox.type = "checkbox";
					newChkBox.id = result.entities[i][self._entityName+"id"];
					newChkBox.name = result.entities[i][self._fieldName];
					newChkBox.value = result.entities[i][self._entityName+"id"];
					if(self._context.parameters.fieldValue.formatted != ""){
						var indexString = self?._context?.parameters?.fieldValue?.formatted?.indexOf(result.entities[i][self._entityName+"id"]);
						if(indexString != -1 && indexString != undefined){
							newChkBox.checked = true;
						}
					}
					newChkBox.addEventListener("change", self._checkBoxChanged);

					newLabel.innerHTML = result.entities[i][self._fieldName];
					newLabel.htmlFor = result.entities[i][self._entityName+"id"];
					
					newUList.appendChild(newChkBox);
					newUList.appendChild(newLabel);
					self._unorderedList.appendChild(newUList);
					
				}                    
				// perform additional operations on retrieved records
			},
			function (error) {
				messageString += error.message;
				self._errorLabel.innerHTML = messageString;
				console.log(error.message);
				// handle error conditions
			}
		);	
		this._mainContainer.appendChild(this._unorderedList);
		this._mainContainer.appendChild(this._errorLabel);
		this._container.appendChild(this._mainContainer);
	}

	
	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void
	{
		// Add code to update control view
		this._context = context;
		if(this._context != null){
			var multiSelectAttribute = this?._context?.parameters?.fieldValue?.attributes?.LogicalName;
			var temp = this._context.parameters.fieldValue.formatted ?? "";
			this._guidList = temp.split(',');
		}
	}

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs
	{
		return {
			fieldValue : this._guidList.join().replace(/^,/, '')
		};
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void
	{
		// Add code to cleanup control if necessary
	}

	public checkBoxChanged(evnt : Event) : void {
		var targetInput = <HTMLInputElement>evnt.target;
		var guidList = this._guidList;

		if(targetInput.checked)
		{			
			guidList?.push(targetInput.value);
		}
		else{
			var ind = guidList.indexOf(targetInput.value);
			guidList?.splice(ind, 1);
		}
		if(guidList != undefined){			
			this._guidList = guidList.filter(function() { return true; });
		}
		this._notifyOutputChanged();
	}
}