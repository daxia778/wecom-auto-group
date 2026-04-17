export namespace main {
	
	export class AppState {
	    processed_customers: string[];
	    target_userid: string;
	    fixed_members: string[];
	    group_owner: string;
	
	    static createFrom(source: any = {}) {
	        return new AppState(source);
	    }
	
	    constructor(source: any = {}) {
	        if ('string' === typeof source) source = JSON.parse(source);
	        this.processed_customers = source["processed_customers"];
	        this.target_userid = source["target_userid"];
	        this.fixed_members = source["fixed_members"];
	        this.group_owner = source["group_owner"];
	    }
	}
	export class Contact {
	    external_userid: string;
	    name: string;
	    type: number;
	    corp_name: string;
	
	    static createFrom(source: any = {}) {
	        return new Contact(source);
	    }
	
	    constructor(source: any = {}) {
	        if ('string' === typeof source) source = JSON.parse(source);
	        this.external_userid = source["external_userid"];
	        this.name = source["name"];
	        this.type = source["type"];
	        this.corp_name = source["corp_name"];
	    }
	}
	export class GroupChat {
	    chat_id: string;
	    name: string;
	    owner: string;
	    member_count: number;
	
	    static createFrom(source: any = {}) {
	        return new GroupChat(source);
	    }
	
	    constructor(source: any = {}) {
	        if ('string' === typeof source) source = JSON.parse(source);
	        this.chat_id = source["chat_id"];
	        this.name = source["name"];
	        this.owner = source["owner"];
	        this.member_count = source["member_count"];
	    }
	}
	export class Member {
	    userid: string;
	    name: string;
	    status: number;
	
	    static createFrom(source: any = {}) {
	        return new Member(source);
	    }
	
	    constructor(source: any = {}) {
	        if ('string' === typeof source) source = JSON.parse(source);
	        this.userid = source["userid"];
	        this.name = source["name"];
	        this.status = source["status"];
	    }
	}

}

