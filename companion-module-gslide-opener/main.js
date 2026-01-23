const { InstanceBase, runEntrypoint, combineRgb } = require('@companion-module/base')
const UpdateActions = require('./actions.js')

class GoogleSlidesOpenerInstance extends InstanceBase {
	constructor(internal) {
		super(internal)
		console.log('[gslide-opener] Constructor called')
		this.log('info', '=== Constructor called ===')
		
		// State for variables and feedbacks
		this.state = {
			presentationOpen: false,
			notesOpen: false,
			currentSlide: null,
			totalSlides: null,
			presentationUrl: null,
			slideInfo: null,
			isFirstSlide: null,
			isLastSlide: null,
			nextSlide: null,
			previousSlide: null,
			presentationTitle: null,
			timerElapsed: null,
			presentationDisplayId: null,
			notesDisplayId: null,
			loginState: false,
			loggedInUser: null
		}
		
		// Polling interval
		this.pollInterval = null
	}

	async init(config) {
		console.log('[gslide-opener] init() called with config:', JSON.stringify(config))
		this.log('info', '=== init() called ===')
		this.log('info', `Raw config received: ${JSON.stringify(config)}`)
		
		// Merge config with defaults
		this.config = {
			host: '127.0.0.1',
			port: '9595',
			...config
		}

		console.log('[gslide-opener] Merged config:', JSON.stringify(this.config))
		this.log('info', `Merged config: host=${this.config.host}, port=${this.config.port}`)

		// Set initial status
		this.updateStatus('ok', 'Initializing...')
		this.log('info', 'Status set to: Initializing...')

		// Update actions
		this.log('info', 'Calling updateActions()...')
		this.updateActions()
		this.log('info', 'Actions updated')

		// Set up variables
		this.log('info', 'Setting up variables...')
		this.setupVariables()
		this.log('info', 'Variables set up')

		// Set up feedbacks
		this.log('info', 'Setting up feedbacks...')
		this.setupFeedbacks()
		this.log('info', 'Feedbacks set up')

		// Test connection after module is ready
		this.log('info', 'Starting connection test...')
		this.testConnection().catch(error => {
			this.log('warn', `Connection test failed during init: ${error.message}`)
		})

		// Start polling for state updates
		this.startPolling()

		this.log('info', '=== Module initialization completed ===')
		console.log('[gslide-opener] init() completed')
	}

	async destroy() {
		console.log('[gslide-opener] destroy() called')
		this.log('info', '=== destroy() called ===')
		
		// Stop polling
		if (this.pollInterval) {
			clearInterval(this.pollInterval)
			this.pollInterval = null
			this.log('info', 'Polling stopped')
		}
	}

	async configUpdated(config) {
		console.log('[gslide-opener] configUpdated() called with:', JSON.stringify(config))
		this.log('info', '=== configUpdated() called ===')
		this.log('info', `New config received: ${JSON.stringify(config)}`)
		
		// Merge config with defaults
		this.config = {
			host: '127.0.0.1',
			port: '9595',
			...config
		}
		
		console.log('[gslide-opener] Merged updated config:', JSON.stringify(this.config))
		this.log('info', `Updated merged config: host=${this.config.host}, port=${this.config.port}`)
		await this.testConnection()
	}

	getConfigFields() {
		console.log('[gslide-opener] getConfigFields() called')
		this.log('info', '=== getConfigFields() called ===')
		
		const fields = [
			{
				type: 'static-text',
				id: 'info',
				width: 12,
				label: 'Information',
				value: 'This module controls the Google Slides Opener Electron app. Make sure the app is running on the same computer.',
			},
			{
				type: 'textinput',
				id: 'host',
				label: 'Host',
				width: 6,
				default: '127.0.0.1',
			},
			{
				type: 'number',
				id: 'port',
				label: 'Port',
				width: 6,
				min: 1,
				max: 65535,
				default: 9595,
			},
		]
		
		console.log('[gslide-opener] getConfigFields() returning:', JSON.stringify(fields))
		this.log('info', `Returning ${fields.length} config fields`)
		return fields
	}

	updateActions() {
		console.log('[gslide-opener] updateActions() called')
		this.log('info', '=== updateActions() called ===')
		UpdateActions(this)
		this.log('info', 'Actions updated successfully')
	}

	// Test connection to the API
	async testConnection() {
		console.log('[gslide-opener] testConnection() called')
		this.log('info', '=== Testing connection ===')
		this.log('info', `Connecting to ${this.config.host}:${this.config.port}`)
		try {
			const response = await this.apiRequest('GET', '/api/status')
			console.log('[gslide-opener] Connection successful:', response)
			this.log('info', 'Connected to Google Slides Opener')
			this.log('info', `Response: ${JSON.stringify(response)}`)
			this.updateStatus('ok', 'Connected')
			
			// Update state immediately on successful connection
			await this.updateState()
		} catch (error) {
			console.log('[gslide-opener] Connection failed:', error)
			this.log('error', `Failed to connect to Google Slides Opener: ${error.message}`)
			this.updateStatus('error', `Connection failed: ${error.message}`)
		}
	}

	// Make API requests
	async apiRequest(method, endpoint, data = null) {
		const http = require('http')

		console.log(`[gslide-opener] API Request: ${method} ${endpoint}`)
		this.log('debug', `API Request: ${method} ${endpoint}`)

		return new Promise((resolve, reject) => {
			const options = {
				hostname: this.config.host || '127.0.0.1',
				port: this.config.port || 9595,
				path: endpoint,
				method: method,
				headers: {
					'Content-Type': 'application/json',
				},
				timeout: 5000,
			}
			
			console.log('[gslide-opener] Request options:', JSON.stringify(options))

			const req = http.request(options, (res) => {
				let responseData = ''
				res.on('data', (chunk) => {
					responseData += chunk
				})
				res.on('end', () => {
					try {
						const response = JSON.parse(responseData)
						if (res.statusCode === 200) {
							resolve(response)
						} else {
							reject(new Error(response.error || 'Request failed'))
						}
					} catch (error) {
						reject(error)
					}
				})
			})

			req.on('error', (error) => {
				reject(error)
			})

			req.on('timeout', () => {
				req.destroy()
				reject(new Error('Request timeout'))
			})

			if (data) {
				req.write(JSON.stringify(data))
			}
			req.end()
		})
	}

	// Set up variable definitions
	setupVariables() {
		const variables = [
			{
				variableId: 'presentation_open',
				name: 'Presentation Open'
			},
			{
				variableId: 'notes_open',
				name: 'Speaker Notes Open'
			},
			{
				variableId: 'current_slide',
				name: 'Current Slide Number'
			},
			{
				variableId: 'total_slides',
				name: 'Total Slides'
			},
			{
				variableId: 'slide_info',
				name: 'Slide Info (e.g. "3 / 10")'
			},
			{
				variableId: 'next_slide',
				name: 'Next Slide Number'
			},
			{
				variableId: 'previous_slide',
				name: 'Previous Slide Number'
			},
			{
				variableId: 'is_first_slide',
				name: 'Is First Slide'
			},
			{
				variableId: 'is_last_slide',
				name: 'Is Last Slide'
			},
			{
				variableId: 'presentation_url',
				name: 'Presentation URL'
			},
			{
				variableId: 'presentation_title',
				name: 'Presentation Title'
			},
			{
				variableId: 'timer_elapsed',
				name: 'Timer Elapsed (e.g. "00:00:06")'
			},
			{
				variableId: 'presentation_display_id',
				name: 'Presentation Display ID'
			},
			{
				variableId: 'notes_display_id',
				name: 'Notes Display ID'
			},
			{
				variableId: 'login_state',
				name: 'Login State (Yes/No)'
			},
			{
				variableId: 'logged_in_user',
				name: 'Logged In User (Email)'
			}
		]
		
		this.setVariableDefinitions(variables)
		this.log('info', `Defined ${variables.length} variables`)
	}

	// Set up feedback definitions
	setupFeedbacks() {
		const feedbacks = {
			presentation_open: {
				type: 'boolean',
				name: 'Presentation is Open',
				description: 'Indicates when a presentation is currently open',
				defaultStyle: {
					color: combineRgb(255, 255, 255),
					bgcolor: combineRgb(0, 200, 0)
				},
				options: [],
				callback: (feedback) => {
					return this.state.presentationOpen === true
				},
				showInvert: true
			},
			notes_open: {
				type: 'boolean',
				name: 'Speaker Notes are Open',
				description: 'Indicates when the speaker notes window is open',
				defaultStyle: {
					color: combineRgb(255, 255, 255),
					bgcolor: combineRgb(0, 150, 255)
				},
				options: [],
				callback: (feedback) => {
					return this.state.notesOpen === true
				},
				showInvert: true
			},
			on_slide: {
				type: 'boolean',
				name: 'On Specific Slide',
				description: 'Indicates when the presentation is on a specific slide number',
				defaultStyle: {
					color: combineRgb(255, 255, 255),
					bgcolor: combineRgb(255, 150, 0)
				},
				options: [
					{
						type: 'number',
						id: 'slide',
						label: 'Slide Number',
						min: 1,
						default: 1
					}
				],
				callback: (feedback) => {
					const targetSlide = feedback.options.slide
					return this.state.currentSlide !== null && this.state.currentSlide === targetSlide
				},
				showInvert: true
			},
			is_first_slide: {
				type: 'boolean',
				name: 'Is First Slide',
				description: 'Indicates when the presentation is on the first slide',
				defaultStyle: {
					color: combineRgb(255, 255, 255),
					bgcolor: combineRgb(100, 200, 100)
				},
				options: [],
				callback: (feedback) => {
					return this.state.isFirstSlide === true
				},
				showInvert: true
			},
			is_last_slide: {
				type: 'boolean',
				name: 'Is Last Slide',
				description: 'Indicates when the presentation is on the last slide',
				defaultStyle: {
					color: combineRgb(255, 255, 255),
					bgcolor: combineRgb(200, 100, 100)
				},
				options: [],
				callback: (feedback) => {
					return this.state.isLastSlide === true
				},
				showInvert: true
			},
			login_state: {
				type: 'boolean',
				name: 'Logged In to Google',
				description: 'Indicates when logged in to Google account',
				defaultStyle: {
					color: combineRgb(255, 255, 255),
					bgcolor: combineRgb(66, 133, 244)
				},
				options: [],
				callback: (feedback) => {
					return this.state.loginState === true
				},
				showInvert: true
			}
		}
		
		this.setFeedbackDefinitions(feedbacks)
		this.log('info', `Defined ${Object.keys(feedbacks).length} feedbacks`)
	}

	// Start polling for state updates
	startPolling() {
		// Poll immediately, then every 1 second
		this.updateState()
		this.pollInterval = setInterval(() => {
			this.updateState()
		}, 1000)
		this.log('info', 'Started polling for state updates (1s interval)')
	}

	// Update state from API and refresh variables/feedbacks
	async updateState() {
		try {
			const response = await this.apiRequest('GET', '/api/status')
			
		// Update internal state with all available fields
		const newState = {
			presentationOpen: response.presentationOpen === true,
			notesOpen: response.notesOpen === true,
			currentSlide: response.currentSlide !== null && response.currentSlide !== undefined ? response.currentSlide : null,
			totalSlides: response.totalSlides !== null && response.totalSlides !== undefined ? response.totalSlides : null,
			presentationUrl: response.presentationUrl || null,
			slideInfo: response.slideInfo || null,
			isFirstSlide: response.isFirstSlide === true,
			isLastSlide: response.isLastSlide === true,
			nextSlide: response.nextSlide !== null && response.nextSlide !== undefined ? response.nextSlide : null,
			previousSlide: response.previousSlide !== null && response.previousSlide !== undefined ? response.previousSlide : null,
			presentationTitle: response.presentationTitle || null,
			timerElapsed: response.timerElapsed || null,
			presentationDisplayId: response.presentationDisplayId !== null && response.presentationDisplayId !== undefined ? response.presentationDisplayId : null,
			notesDisplayId: response.notesDisplayId !== null && response.notesDisplayId !== undefined ? response.notesDisplayId : null,
			loginState: response.loginState === true,
			loggedInUser: response.loggedInUser || null
		}
		
		// Check if state changed (compare all fields)
		const stateChanged = 
			this.state.presentationOpen !== newState.presentationOpen ||
			this.state.notesOpen !== newState.notesOpen ||
			this.state.currentSlide !== newState.currentSlide ||
			this.state.totalSlides !== newState.totalSlides ||
			this.state.presentationUrl !== newState.presentationUrl ||
			this.state.slideInfo !== newState.slideInfo ||
			this.state.isFirstSlide !== newState.isFirstSlide ||
			this.state.loginState !== newState.loginState ||
			this.state.loggedInUser !== newState.loggedInUser ||
			this.state.isLastSlide !== newState.isLastSlide ||
			this.state.nextSlide !== newState.nextSlide ||
			this.state.previousSlide !== newState.previousSlide ||
			this.state.presentationTitle !== newState.presentationTitle ||
			this.state.timerElapsed !== newState.timerElapsed ||
			this.state.presentationDisplayId !== newState.presentationDisplayId ||
			this.state.notesDisplayId !== newState.notesDisplayId
		
		if (stateChanged) {
			this.state = newState
			
			// Update all variables
			this.setVariableValues({
				presentation_open: this.state.presentationOpen ? 'Yes' : 'No',
				notes_open: this.state.notesOpen ? 'Yes' : 'No',
				current_slide: this.state.currentSlide !== null ? String(this.state.currentSlide) : '',
				total_slides: this.state.totalSlides !== null ? String(this.state.totalSlides) : '',
				slide_info: this.state.slideInfo || '',
				next_slide: this.state.nextSlide !== null ? String(this.state.nextSlide) : '',
				previous_slide: this.state.previousSlide !== null ? String(this.state.previousSlide) : '',
				is_first_slide: this.state.isFirstSlide ? 'Yes' : 'No',
				is_last_slide: this.state.isLastSlide ? 'Yes' : 'No',
				presentation_url: this.state.presentationUrl || '',
				presentation_title: this.state.presentationTitle || '',
				timer_elapsed: this.state.timerElapsed || '',
				presentation_display_id: this.state.presentationDisplayId !== null ? String(this.state.presentationDisplayId) : '',
				notes_display_id: this.state.notesDisplayId !== null ? String(this.state.notesDisplayId) : '',
				login_state: this.state.loginState ? 'Yes' : 'No',
				logged_in_user: this.state.loggedInUser || ''
			})
			
			// Trigger feedback updates
			this.checkFeedbacks('presentation_open', 'notes_open', 'on_slide', 'is_first_slide', 'is_last_slide', 'login_state')
			
			this.log('debug', `State updated: presentation=${this.state.presentationOpen}, notes=${this.state.notesOpen}, slide=${this.state.currentSlide}/${this.state.totalSlides}, title=${this.state.presentationTitle || 'N/A'}`)
		}
		} catch (error) {
			// Silently fail - connection might be down, don't spam logs
			// Only log if we previously had a connection
			if (this.state.presentationOpen || this.state.notesOpen) {
				this.log('debug', `Failed to update state: ${error.message}`)
			}
		}
	}
}

console.log('[gslide-opener] Module loaded, calling runEntrypoint...')
runEntrypoint(GoogleSlidesOpenerInstance, [])
console.log('[gslide-opener] runEntrypoint called')
