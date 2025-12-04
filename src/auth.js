require('dotenv').config();

class OmniHRAuth {
	constructor() {
		this.baseURL =
			process.env.OMNIHR_BASE_URL || 'https://api.omnihr.co/api/v1';
		this.username = process.env.OMNIHR_USERNAME;
		this.password = process.env.OMNIHR_PASSWORD;
		this.subdomain = process.env.OMNIHR_SUBDOMAIN;
		this.token = null;

		if (!this.username || !this.password) {
			throw new Error(
				'OMNIHR_USERNAME and OMNIHR_PASSWORD environment variables are required'
			);
		}

		if (!this.subdomain) {
			throw new Error('OMNIHR_SUBDOMAIN environment variable is required');
		}
	}

	async login() {
		const loginUrl = `${this.baseURL}/auth/token/`;

		const response = await fetch(loginUrl, {
			method: 'POST',
			headers: {
				'Content-Type': 'application/x-www-form-urlencoded',
				'x-subdomain': this.subdomain,
			},
			body: new URLSearchParams({
				username: this.username,
				password: this.password,
			}).toString(),
		});

		const responseText = await response.text();

		if (!response.ok) {
			throw new Error(
				`Login failed with status ${response.status}: ${responseText}`
			);
		}

		const data = JSON.parse(responseText);

		this.token = data.access || data.token || data.access_token;

		if (!this.token) {
			throw new Error('No token found in response');
		}

		return this.token;
	}

	async getToken() {
		if (!this.token) {
			await this.login();
		}
		return this.token;
	}

	getSubdomain() {
		return this.subdomain;
	}

	async getAuthHeaders() {
		const token = await this.getToken();
		return {
			Authorization: `Bearer ${token}`,
			'x-subdomain': this.subdomain,
			'Content-Type': 'application/json',
		};
	}
}

module.exports = OmniHRAuth;
